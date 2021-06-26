using System;
using OpenMcdf;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using OpenMcdf.Extensions;
using OpenMcdf.Extensions.OLEProperties;

namespace MsiReader
{
    static class Win32Error
    {
        public const int NO_ERROR = 0;
        public const int ERROR_NO_MORE_ITEMS = 259;

    }
    class MsiPull
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

        public static int DrawFromMsi(String fileName)
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
                    Console.WriteLine(buffer.ToString());
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
    public class Class1
    {
        static void Main(string[] args)
        {
            String filename = "Setupdistex.msi";
            CompoundFile cf = new CompoundFile(filename);
            //  Console.WriteLine(cf.ToString());
            // Console.WriteLine(cf.GetType().ToString());
            //for (int i = 0; i <cf.GetNumDirectories(); ++i)
            //{
            //    Console.WriteLine(cf.GetNameDirEntry(i));
                
            //}
            CFStream fStream = cf.RootStorage.GetStream("\u0005SummaryInformation");
            //byte[] temp = fStream.GetData();
            //var data = temp.GetEnumerator();
            //while(data.MoveNext())
            //{
            //    Console.Write(Convert.ToChar(data.Current));
                
            //}



            // CompoundFile cf = new CompoundFile(filename);
            //  CFStream foundStream = cf.RootStorage.GetStream("\u0005SummaryInformation");
            var container = fStream.AsOLEPropertiesContainer();
            foreach (var property in container.Properties)
                Console.WriteLine($"{property.PropertyName}: {property.Value}");

            //Console.WriteLine(cf.RootStorage.);

            // Console.WriteLine(fStream.ToString());


            //TODO:: nađi način za izvući podatke summary infoa
            //OLEPropertiesContainer.SummaryInfoProperties summaryInfo;
            MsiPull.DrawFromMsi("Setupdistex.msi");

            // OLEPropertiesContainer container = fStream.AsOLEPropertiesContainer();
            // OLEPropertiesContainer.SummaryInfoProperties summaryInfo = new OLEPropertiesContainer.SummaryInfoProperties();

            /*Nisam siguran na koji bi način sad iz msi-a izvukao SummaryInfo
             cf.ToString() mi ne vraća ništa pa nisam siguran je li file učitan kako treba. 

            Program ne crasha, ali na GetStream dobijem "Cannot find item [Summary Information] within the current storage"
            Probao sam sa "_" umjesto razmaka i provjerom slova više puta, nisam siguran radim li ulazak u file kako treba
            zbog ovog ToStringa.

            cf.GetType().ToString(); također ništa ne daje
            Što se tiče ovog containera mislim da će mi negdje trebati, ili nekako ću koristiti ovaj SummaryInfoProperties 
             */
            cf.Close();
            // container.
        }
        }
    }
