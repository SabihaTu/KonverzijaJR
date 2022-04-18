using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KonverzijaExcXml
{
    class Program
    {
       
        static void Main(string[] args)
        {
            int i ;
            string direktorij = "C:/SFTP";//"c:/JR/import";       
            DateTime localDate = DateTime.Now;
            string format = localDate.ToString("yyyyMMddHHmmss");
            String cultureName = "de-DE";
            var culture = new CultureInfo(cultureName);
            //  Kreiraj objekat obrada
            Obrada obrada = new Obrada();

          
            //
            //Ucitaj parametre baze
            //
            ParametriXml parametri = new ParametriXml();
            parametri.putConfig = direktorij;
            parametri.format = format;
            parametri.citajXml();
            //  Zapamti eMail administratora
            if (parametri.mailAdmin.Equals(""))
                {
                    //                obrada.listeMailAdmin.Add("zbakir@comp-it.ba");
                    obrada.addEMail(obrada.listeMailAdmin, "zbakir@comp-it.ba", " - Administrator");
                }
            else
                {
                    //                obrada.listeMailAdmin.Add(parametri.mailAdmin);
                    obrada.addEMail(obrada.listeMailAdmin, parametri.mailAdmin, " - Administrator");
                }

            direktorij = parametri.putLog;
            var fileTrace = Path.Combine(direktorij, "Trace" + format + ".txt");
            var fileGreske = Path.Combine(direktorij, "Obavijesti" + format + ".txt");
            Encoding encoding = Encoding.Unicode;
            StreamWriter fs = new StreamWriter(fileTrace, false, Encoding.Unicode);
            StreamWriter fs1 = new StreamWriter(fileGreske, false, Encoding.Unicode);
            //  Pisi xml parametre u log datoteku
            parametri.pisiDBParametreToFs(fs);
            //  Oznaci start programa u log datoteku
            fs.WriteLine("IMPORT - start " + localDate.ToString(culture) + " " + localDate.Kind);
            fs.Flush();
            //konekcija na bazu
            obrada.podaciInit(fs,parametri);
            direktorij = parametri.putanja;
           // obrada excel fajla
           FileInfo[] fajlovi = pokupiFajlove(direktorij);

            // petlja za excel fajlove
            if (fajlovi.Length >= 1)
            { 
                for ( i = 0; i < 1; i++)
            
            {
               // if (i < 1)
                
                    try
                    {
                        obrada.listeMail.Clear();
                        // ime excel fajla
                        String ime = fajlovi[i].FullName;
                        String samoIme = fajlovi[i].Name;
                        // citanje i kontrola isplata, slanje gresaka na mail ukoliko postoje
                        int postojiGreska = obrada.provjeraExcelDatoteke(ime, fs, fs1, parametri, ime);
                        i++;
                        // sad treba ponovno citanje isplata i upis u bazu ukoliko u 1 dokumentu nema gresaka
                        if (postojiGreska == 0)
                        {
                            //INSERT u BAZU
                            obrada.unesiIsplateUBazu(fs1, samoIme);
                        }

                        if (i != fajlovi.Length)
                        {

                            // otvaranje sljedeceg fajla sa greskama
                            direktorij = parametri.putLog;
                            localDate = DateTime.Now;
                            format = localDate.ToString("yyyyMMddHHmmss");
                            fileGreske = Path.Combine(direktorij, "Obavijesti" + format + ".txt");
                            fs1 = new StreamWriter(fileGreske, false, Encoding.Unicode);
                            fs1.WriteLine(localDate.ToString(culture) + " " + localDate.Kind);
                            fs1.Flush();
                        }
                    }
                    catch (Exception e)
                    {
                        fs.WriteLine("Greska - petlja za excel fajlove " + e.Message);
                    }
                }

            }
        }

        public static FileInfo[] pokupiFajlove(String folder)
        {
            DirectoryInfo di = new DirectoryInfo(folder);
            FileInfo[] files = di.GetFiles("JR*.xlsx");
            return files;
        }
    }
}
