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
        static extern uint MsiRecordReadStream(IntPtr hRecord, uint iField,[Out] byte[] szDataBuf, ref int pcbDataBuf);

        [DllImport("msi.dll", ExactSpelling = true)]
        static extern int MsiRecordDataSize(IntPtr hRecord, int iField);

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
        public static string GetItemData(String fullPathName,ref List<String>dataString,uint indexOfItem)
        {
            try
            {
                if (MsiOpenDatabase(fullPathName, IntPtr.Zero, out IntPtr hDatabase) != Win32Error.NO_ERROR)
                {
                    Console.WriteLine(fullPathName);
                    return "Failed to open database";
                }

                if (MsiDatabaseOpenView(hDatabase, "SELECT `Name` FROM _Tables", out IntPtr hView) != Win32Error.NO_ERROR)
                {
                    return "Failed to open view";
                }

                MsiViewExecute(hView, IntPtr.Zero);

                while (MsiViewFetch(hView, out IntPtr hRecord) != Win32Error.ERROR_NO_MORE_ITEMS)
                {
                    //List<String> allFileNames = new List<String>();
                    //MsiPull.DrawFromMsi(fullPathName, ref allFileNames);
                    //String data=allFileNames.Find(x => x == fullPathName);
                    //return data;

                    //int capacity = dataBuf.Length;
                    int capacity = MsiRecordDataSize(hRecord,(int) indexOfItem+1);
                    var dataBuf = new byte[capacity];
                    //for (int i = 0; i < dataBuf.Length; i++)
                    //{
                    //    dataBuf[i] = 0x20;
                    //    String toAdd = Environment.NewLine+dataBuf[i].ToString();
                    //    dataString.Add(toAdd);
                    //}
                    if (MsiRecordReadStream(hRecord, indexOfItem+1, dataBuf, ref capacity) != Win32Error.NO_ERROR)
                    {
                        return "Failed to read stream";
                    }
                    dataString.Add(dataBuf.ToString());
                }
                return "No fail";
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return "Exception";
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
