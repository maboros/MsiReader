
using System;
using OpenMcdf;
using System.IO;
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
        public const int MSI_NULL_INTEGER = unchecked((int)0x80000000);

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
        public static int DrawFromMsi(String fileName,ref List<String> returnNames)
        {
            try
            {
                if (Msi.OpenDatabase(fileName, IntPtr.Zero, out IntPtr hDatabase) != Win32Error.NO_ERROR)
                {
                    Console.WriteLine(fileName);
                    Console.WriteLine("Failed to open dabase");
                    return 1;
                }

                if (Msi.DatabaseOpenView(hDatabase, "SELECT `Name` FROM _Tables", out IntPtr hView) != Win32Error.NO_ERROR)
                {
                    Console.WriteLine("Failed to open view");
                    return 1;
                }

                Msi.ViewExecute(hView, IntPtr.Zero);

                while (Msi.ViewFetch(hView, out IntPtr hRecord) != Win32Error.ERROR_NO_MORE_ITEMS)
                {
                    StringBuilder buffer = new StringBuilder(256);
                    int capacity = buffer.Capacity;
                    if (Msi.RecordGetString(hRecord, 1, buffer, ref capacity) != Win32Error.NO_ERROR)
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
                if (Msi.OpenDatabase(fullPathName, IntPtr.Zero, out IntPtr hDatabase) != Win32Error.NO_ERROR)
                {
                    Console.WriteLine(fullPathName);
                    return 1;
                }

                if (Msi.DatabaseOpenView(hDatabase, $"SELECT * FROM `{tableName}`", out IntPtr hView) != Win32Error.NO_ERROR)
                {
                    return 2;
                }
                Msi.ViewExecute(hView, IntPtr.Zero);
                while (Msi.ViewFetch(hView, out IntPtr hRecord) != Win32Error.ERROR_NO_MORE_ITEMS)
                {
                    // dohvaća informacije o imenima stupaca; ako je drugi argument 1, vratit će 
                    // informacije o tipu u pojedinim stupcima (https://docs.microsoft.com/en-us/windows/win32/msi/column-definition-format)
                    Msi.ViewGetColumnInfo(hView, 1, out IntPtr hColumnRecord);
                    // dohvaća broj stupaca
                    var fieldCount = Msi.RecordGetFieldCount(hRecord);
                    // za svaki stupac...
                    for (int i = 0; i < fieldCount; ++i)
                    {
                        // ..dohvaća duljinu imena...
                        var dataSize = Msi.RecordDataSize(hColumnRecord, i + 1);
                        // ...i samo ime
                        StringBuilder buffer = new StringBuilder(dataSize + 1);
                        int capacity = buffer.Capacity;
                        if (Msi.RecordGetString(hColumnRecord, i + 1, buffer, ref capacity) != Win32Error.NO_ERROR)
                        {
                            return 3;
                        }
                        if (buffer.ToString().ToLower().Equals("i2") || buffer.ToString().ToLower().Equals("i4"))
                        {
                            int num = Msi.RecordGetInteger(hRecord, i + 1);
                            if (num == Win32Error.MSI_NULL_INTEGER)
                            {
                                return 4;
                            }
                            dataList.Add(num.ToString());
                        }
                        else if (buffer.ToString().ToLower().Equals("v0"))
                        {
                            Object obj;
                        }
                        else
                        {
                            StringBuilder dataStr=new StringBuilder(dataSize+1);
                            int cap = dataStr.Capacity;
                            if(Msi.RecordGetString(hColumnRecord, i + 1, dataStr, ref cap) != Win32Error.NO_ERROR)
                            {
                                return 6;
                            }
                            dataList.Add(dataStr.ToString());
                        }

                        //if (i == 0) { dataList.Add(buffer.ToString()); }
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
