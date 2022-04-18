using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KonverzijaExcXml
{
    public class Baza
    {
        SqlConnection connection;
        public string greskaTekst;
        public List<string> listGresaka = new List<string>();
        public bool greska, greskaKonv, greskaZaPrekid, greskaPostoji;
        public SqlConnection createDBConnection(StreamWriter fs, ParametriXml par)
        {
            SqlConnection conn;
            try
            {
                conn = new SqlConnection(getKonekcioniString(par));
            }
            catch (SqlException e)
            {
                greskaTekst = "Greška konekcije: " + e.Message;
                upisiGresku(listGresaka, greskaTekst);
                fs.WriteLine(greskaTekst);
                greska = true;
                conn = connection;
                fs.Flush();
            }
            return conn;
        }
        public string getKonekcioniString(ParametriXml par)
        {
            string konString;
            if (par.korisnikDb.ToLower().Equals("trusted_connection"))
            {
                konString = "Server = " + par.urlDb + "; Initial Catalog = " + par.bazaDb + "; Trusted_Connection = True ";
            }
            else
            {
                konString = "Data Source = " + par.urlDb + "; Initial Catalog = " + par.bazaDb + "; User ID = " + par.korisnikDb + "; Password =" + par.lozinkaDb;
            }
            return konString;
        }

        public void createConnection(StreamWriter fs, ParametriXml par)
        {
            connection = createDBConnection(fs, par);
        }

        public void upisiGresku(List<string> listGresaka, string poruka)
        {
            if (poruka == "")
            {
                return;
            }
            foreach (string s in listGresaka)
            {
                if (s == poruka)
                {
                    return;
                }
            }

            listGresaka.Add(poruka);
            return;
        }
    }
}
