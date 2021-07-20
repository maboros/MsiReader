using System;
using OpenMcdf;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using OpenMcdf.Extensions;
using OpenMcdf.Extensions.OLEProperties;
using System.Collections.Generic;

namespace MsiReader
{
    static class Win32Error
    {
        public const int NO_ERROR = 0;
        public const int ERROR_NO_MORE_ITEMS = 259;

    }
    public class SummaryInfoProps
    {
        public String name;
        public String value;
        public SummaryInfoProps(String name, String value)
        {
            this.name = name;
            this.value = value;
        }
    }
    public class MsiPull
    {
        [DllImport("msi.dll", SetLastError = true)]
        static extern uint MsiOpenDatabase(string szDatabasePath, IntPtr phPersist, out IntPtr phDatabase);

        [DllImport("msi.dll", CharSet = CharSet.Unicode)]
        static extern int MsiDatabaseOpenView(IntPtr hDatabase, [MarshalAs(UnmanagedType.LPWStr)] string szQuery, out IntPtr phView);

        [DllImport("msi.dll", CharSet = CharSet.Unicode)]
        static extern int MsiViewExecute(IntPtr hView, IntPtr hRecord);

        [DllImport("msi.dll", CharSet = CharSet.Unicode)]
        static extern uint MsiViewFetch(IntPtr hView, out IntPtr hRecord);


        [DllImport("msi.dll", CharSet = CharSet.Unicode)]
        static extern int MsiRecordGetString(IntPtr hRecord, int iField, StringBuilder szValueBuf, ref int pcchValueBuf);

        [DllImport("msi.dll", ExactSpelling = true)]
        static extern int MsiRecordDataSize(IntPtr hRecord, int iField);

        const int MSICOLINFO_NAMES = 0;  // return column names
        const int MSICOLINFO_TYPES = 1;  // return column definitions, datatype code followed by width
        [DllImport("msi.dll", ExactSpelling = true)]
        static extern uint MsiViewGetColumnInfo(IntPtr hView, int eColumnInfo, out IntPtr hRecord);

        [DllImport("msi.dll", ExactSpelling = true)]
        static extern int MsiRecordGetFieldCount(IntPtr hRecord);

        public static int DrawFromMsi(String fileName,ref List<String> returnNames)
        {
            try
            {
                if (MsiOpenDatabase(fileName, IntPtr.Zero, out IntPtr hDatabase) != Win32Error.NO_ERROR)
                {
                    Console.WriteLine(fileName);
                    Console.WriteLine("Failed to open dabase");
                    return 1;
                }

                if (MsiDatabaseOpenView(hDatabase, "SELECT `Name` FROM _Tables", out IntPtr hView) != Win32Error.NO_ERROR)
                {
                    Console.WriteLine("Failed to open view");
                    return 1;
                }

                MsiViewExecute(hView, IntPtr.Zero);

                while (MsiViewFetch(hView, out IntPtr hRecord) != Win32Error.ERROR_NO_MORE_ITEMS)
                {
                    StringBuilder buffer = new StringBuilder(256);
                    int capacity = buffer.Capacity;
                    if (MsiRecordGetString(hRecord, 1, buffer, ref capacity) != Win32Error.NO_ERROR)
                    {
                        Console.WriteLine("Failed to get record string");
                        return 1;
                    }
                    returnNames.Add(buffer.ToString());
                }
                return 0;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return 1;
            }
        }
        public static int GetItemData(String fullPathName,ref List<String>dataString,String tableName,ref List<String> dataList)
        {
            try
            {
                if (MsiOpenDatabase(fullPathName, IntPtr.Zero, out IntPtr hDatabase) != Win32Error.NO_ERROR)
                {
                    Console.WriteLine(fullPathName);
                    return 1;
                }

                if (MsiDatabaseOpenView(hDatabase, $"SELECT * FROM `{tableName}`", out IntPtr hView) != Win32Error.NO_ERROR)
                {
                    return 2;
                }
                MsiViewExecute(hView, IntPtr.Zero);
                while (MsiViewFetch(hView, out IntPtr hRecord) != Win32Error.ERROR_NO_MORE_ITEMS)
                {
                    // dohvaća informacije o imenima stupaca; ako je drugi argument 1, vratit će 
                    // informacije o tipu u pojedinim stupcima (https://docs.microsoft.com/en-us/windows/win32/msi/column-definition-format)
                    MsiViewGetColumnInfo(hView, 0, out IntPtr hColumnRecord);
                    // dohvaća broj stupaca
                    var fieldCount = MsiRecordGetFieldCount(hRecord);
                    // za svaki stupac...
                    for (int i = 0; i < fieldCount; ++i)
                    {
                        // ..dohvaća duljinu imena...
                        var dataSize = MsiRecordDataSize(hColumnRecord, i + 1);
                        // ...i samo ime
                        StringBuilder buffer = new StringBuilder(dataSize + 1);
                        int capacity = buffer.Capacity;
                        if (MsiRecordGetString(hColumnRecord, i + 1, buffer, ref capacity) != Win32Error.NO_ERROR)
                        {
                            return 3;
                        }
                        dataList.Add(buffer.ToString());
                    }
                   

                }
                return 0;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return 4;
            }
        }
        public static List<SummaryInfoProps> getSummaryInformation(String fileName)
        {
            CompoundFile cf = new CompoundFile(fileName);
            CFStream fStream = cf.RootStorage.GetStream("\u0005SummaryInformation");
            List < SummaryInfoProps >fullList= new List<SummaryInfoProps>();
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            var container = fStream.AsOLEPropertiesContainer();
            foreach (var property in container.Properties)
            {
                fullList.Add(new SummaryInfoProps(property.PropertyName.ToString(),property.Value.ToString()));
            }
            cf.Close();
            return fullList;
        }
    }   
}
