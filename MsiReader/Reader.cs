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

        public static int DrawFromMsi(String fileName,ref List<String> returnNames)
        {
            try
            {
                if (MsiOpenDatabase(fileName, IntPtr.Zero, out IntPtr hDatabase) != Win32Error.NO_ERROR)
                {
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


    }
    public class Reader
    {
        static void Main(string[] args)
        {
            String fileName = "Setupdistex.msi";
            
            List<String> allFileNames = new List<String>();
            MsiPull.DrawFromMsi(fileName, ref allFileNames);

            foreach (var name in allFileNames)
            {
                Console.WriteLine(name.ToString());
            }


            CompoundFile cf = new CompoundFile(fileName);
            CFStream fStream = cf.RootStorage.GetStream("\u0005SummaryInformation");

            var container = fStream.AsOLEPropertiesContainer();
            foreach (var property in container.Properties)
                Console.WriteLine($"{property.PropertyName}: {property.Value}");

            cf.Close();
        }
        }
    }
