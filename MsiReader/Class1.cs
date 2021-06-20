using System;
using OpenMcdf;
using System.IO;
using OpenMcdf.Extensions;
using OpenMcdf.Extensions.OLEProperties;

namespace MsiReader
{
    public class Class1
    {
        static void Main(string[] args)
        {
            String filename = "Setupdistex.msi";
            CompoundFile cf = new CompoundFile(filename);
            //  Console.WriteLine(cf.ToString());
            // Console.WriteLine(cf.GetType().ToString());
            for (int i = 0; i <cf.GetNumDirectories(); ++i)
            {
                Console.WriteLine(cf.GetNameDirEntry(i));
            }
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
            OLEPropertiesContainer.SummaryInfoProperties summaryInfo;

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
