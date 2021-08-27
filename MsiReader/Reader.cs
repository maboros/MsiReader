
using System;
using OpenMcdf;
using System.Text;
using OpenMcdf.Extensions;
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
    public class Reader
    {
        public static int DrawFromMsi(String fileName,ref List<String> returnNames)
        {
            try
            {
                if (Msi.OpenDatabase(fileName, IntPtr.Zero, out IntPtr hDatabase) != Win32Error.NO_ERROR)
                {
                    Console.WriteLine("Failed to open dabase");
                    return 1;
                }

                if (Msi.DatabaseOpenView(hDatabase, "SELECT `Name` FROM _Tables", out IntPtr hView) != Win32Error.NO_ERROR)
                {
                    Console.WriteLine("Failed to open view");
                    return 2;
                }
                Msi.ViewExecute(hView, IntPtr.Zero);
                while (Msi.ViewFetch(hView, out IntPtr hRecord) != Win32Error.ERROR_NO_MORE_ITEMS)
                {
                    StringBuilder buffer = new StringBuilder(256);
                    int capacity = buffer.Capacity;
                    if (Msi.RecordGetString(hRecord, 1, buffer, ref capacity) != Win32Error.NO_ERROR)
                    {
                        Console.WriteLine("Failed to get record string");
                        return 3;
                    }
                    returnNames.Add(buffer.ToString());
                }
                return 0;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return 4;
            }
        }
        public static int GetItemData(String fullPathName,String tableName,ref List<String> columnList,ref int columnCount,ref List<String> dataList)
        {
            try
            {
                if (Msi.OpenDatabase(fullPathName, IntPtr.Zero, out IntPtr hDatabase) != Win32Error.NO_ERROR)
                {
                    Console.WriteLine("Failed to open dabase");
                    return 1;
                }
                if (Msi.DatabaseOpenView(hDatabase, $"SELECT * FROM `{tableName}`", out IntPtr hView) != Win32Error.NO_ERROR)
                {
                    Console.WriteLine("Failed to open view");
                    return 2;
                }
                Msi.ViewExecute(hView, IntPtr.Zero);
                int x=GetTableColumnInfo(hView, ref columnCount, ref columnList);
                if (x != 0)
                {
                    return x;
                }
                x=GetTableData(hView,ref columnCount,ref dataList);
                return x;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return 5;
            }
        }
        private static int GetTableColumnInfo(IntPtr hView,ref int columnCount,ref List<String> columnList)
        {
            Msi.ViewFetch(hView, out IntPtr hRecord);
            Msi.ViewGetColumnInfo(hView, Msi.MSICOLINFO_NAMES, out IntPtr hColumnRecord);
            columnCount = Msi.RecordGetFieldCount(hRecord);
            if (columnCount <= 0)
            {
                columnCount = 30;
            }
            for (int i = 0; i < columnCount; ++i)
            {
                var dataSize = Msi.RecordDataSize(hColumnRecord, i + 1);
                StringBuilder buffer = new StringBuilder(dataSize + 1);
                int capacity = buffer.Capacity;
                if (Msi.RecordGetString(hColumnRecord, i + 1, buffer, ref capacity) != Win32Error.NO_ERROR)
                {
                    Console.WriteLine("Failed to get record string");
                    return 3;
                }
                if (buffer.ToString() != "")
                {
                    columnList.Add(buffer.ToString());
                }
            }
            return 0;
        }
        private static int GetTableData(IntPtr hView, ref int columnCount, ref List<String> dataList)
        {
            Msi.ViewExecute(hView, IntPtr.Zero);
            while (Msi.ViewFetch(hView, out IntPtr hRecord) != Win32Error.ERROR_NO_MORE_ITEMS)
            {
                Msi.ViewGetColumnInfo(hView, Msi.MSICOLINFO_TYPES, out IntPtr hColumnRecord);
                for (int i = 0; i < columnCount; ++i)
                {
                    var dataSize = Msi.RecordDataSize(hColumnRecord, i + 1);
                    StringBuilder buffer = new StringBuilder(dataSize + 1);
                    int capacity = buffer.Capacity;
                    if (Msi.RecordGetString(hColumnRecord, i + 1, buffer, ref capacity) != Win32Error.NO_ERROR)
                    {
                        Console.WriteLine("Failed to get record string");
                        return 4;
                    }
                    if (buffer.ToString().ToLower().Equals("i2") || buffer.ToString().ToLower().Equals("i4"))
                    {
                        int num = Msi.RecordGetInteger(hRecord, i + 1);
                        if (num == Win32Error.MSI_NULL_INTEGER)
                        {
                            dataList.Add("");
                            continue;
                        }
                        dataList.Add(num.ToString());
                    }
                    else if (buffer.ToString().ToLower().Equals("v0"))
                    {
                        dataList.Add("[Binary data]");
                    }
                    else
                    {
                        StringBuilder dataStr = new StringBuilder(255);
                        int cap = dataStr.Capacity;
                        Msi.RecordGetString(hRecord, i + 1, dataStr, ref cap);
                        dataList.Add(dataStr.ToString());
                    }
                }
            }
            return 0;
        }

        public static List<SummaryInfoProps> GetSummaryInformation(String fileName)
        {
            try
            {
                CompoundFile cf = new CompoundFile(fileName);
                CFStream fStream = cf.RootStorage.GetStream("\u0005SummaryInformation");
                List<SummaryInfoProps> fullList = new List<SummaryInfoProps>();
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                var container = fStream.AsOLEPropertiesContainer();
                foreach (var property in container.Properties)
                {
                    fullList.Add(new SummaryInfoProps(property.PropertyName.ToString(), property.Value.ToString()));
                }
                cf.Close();
                return fullList;
            }
            catch(Exception e)
            {
                List<SummaryInfoProps> fullList = new List<SummaryInfoProps>();
                fullList.Add(new SummaryInfoProps("Error", e.Message.ToString()));
                fullList.Add(new SummaryInfoProps("Advice", "Please make sure the file isn't corrupted and is a .msi file"));
                return fullList;
            }
        }
    }   
}
