using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace KonverzijaExcXml
{
    class Obrada
    {

        public List<string> listUpozorenja = new List<string>();
        public List<string> listeMailAdmin = new List<string>();
        public RegexUtilities ru = new RegexUtilities();
        Microsoft.Office.Interop.Excel.Application xlApp;
        object[,] objectArray;
        Workbook xlWorkbook;
        string odgovor;
        _Worksheet xlWorksheet;
        //_Worksheet xlWorksheet;
        int row;
        //varijable
        public string jib;
        public String emailSlanje;
        public string eMail;
        public string datumIsplate;
        public string brojIsplata;
        public string ukupanIznos;
        public string isDomacinstvo;
        public string vrstaAktaId;
        public string pravniOsnovId;
        public string grupaPravaId;
        public string ucestalostIsplataId;
        public string oblikDpzId;
        public string oblikGrupaDpzId;
        public int oblikGrupaId;
        public string opcinaId;
        public int? idOpcine;
        public string iznos;
        public Double iznosPojedinacan;
        public string iznos1;
        public string iznos2;
        public decimal iznosNum;
        public decimal iznos1Num;
        public decimal iznos2Num;
        public string jmb;
        public string jmbOld = "pocetak";
        public string jibOld = "pocetak";
        public string prezime;
        public string ime;
        public string roditelj;
        public string mjestoRodjenja;
        public string drzavljanstvoId;
        public string posta;
        public string mjesto;
        public string adresa;
        public string pravoId;
        public int rjesenjeId;
        public int predmetId;
        public string protokol;
        public string datumAkta;
        public string odgovornoLice;
        public string unosLice;
        public int oblikId;
        public string nazivInstitucije;
        public string redniBroj;
        public string spol;

        public string clanPrezime1;
        public string clanIme1;
        public string clanSrod1;
        public string clanGodiste1;
        public string vclanLK1;

        public string clanPrezime2;
        public string clanIme2;
        public string clanSrod2;
        public string clanGodiste2;
        public string vclanLK2;

        public string clanPrezime3;
        public string clanIme3;
        public string clanSrod3;
        public string clanGodiste3;
        public string vclanLK3;

        public string clanPrezime4;
        public string clanIme4;
        public string clanSrod4;
        public string clanGodiste4;
        public string vclanLK4;

        public string clanPrezime5;
        public string clanIme5;
        public string clanSrod5;
        public string clanGodiste5;
        public string vclanLK5;

        bool c1;
        SqlConnection connection;
        public string greskaTekst;
        public List<string> listGresaka = new List<string>();
        public bool greska, greskaKonv, greskaZaPrekid, greskaPostoji;

        string sqlKomanda = "";
        DateTime vrijeme;
        List<int> listPravniOsnov = new List<int>();
        List<int> listUcestalostIsplate = new List<int>();
        List<int> listVrstaAkta = new List<int>();
        List<int> listSrodstvo = new List<int>();
        List<int> listDrzavljanstvo = new List<int>();
        List<int> listOpcina = new List<int>();
        List<int> listPrava = new List<int>();
        List<int> listVrstaNaknade = new List<int>();  // grupeprava
        //  Liste sa nepostojecim vrijednostuma FK
        List<int> listPravniOsnovErr = new List<int>();
        List<int> listUcestalostIsplateErr = new List<int>();
        List<int> listVrstaAktaErr = new List<int>();
        List<int> listSrodstvoErr = new List<int>();
        List<int> listDrzavljanstvoErr = new List<int>();
        List<int> listOpcinaErr = new List<int>();
        List<int> listPravaErr = new List<int>();
        List<int> listVrstaNaknadeErr = new List<int>();
        public List<string> listeMail = new List<string>();
        public int k;
        public int brojGresaka;
        public int? uploadId, spisakId, listaId, liceId;
        public System.Globalization.CultureInfo currentCulture = System.Globalization.CultureInfo.InstalledUICulture;
        public System.Globalization.NumberFormatInfo numberFormat;
        string godinaStr, mjesecStr, danStr;
        string justName, jibName;
        public string jibDadoteka;
        int jmbGodina;
        int jmbMjesec;
        int jmbDan;
        int maxDana = 0;
        int maxGodina = DateTime.Now.Year;
        int minGodina = 2016;
        Boolean isPrestupna;
        SqlDataReader reader, readerD;
        public int? firmaId, userId, dpzId;
        public Institucija institucija;

        public Double ukupanIznosStavki;
        public int ukupanBrojStavki;
        String[] listaGresaka;
        DateTime datumIsplate1;

        Boolean ispravna;
        public void podaciInit(StreamWriter fs, ParametriXml parametri)
        {
            brojGresaka = 0;
            numberFormat = (System.Globalization.NumberFormatInfo)currentCulture.NumberFormat.Clone();
            numberFormat.NumberDecimalSeparator = ".";

            DateTime localDate = DateTime.Now;
            godinaStr = localDate.ToString("yyyy");
            mjesecStr = localDate.ToString("MM");
            danStr = localDate.ToString("dd");



            // listeErrClear();


            createConnection(fs, parametri);
            try
            {
                sqlKomanda = "SELECT Id FROM dbo.PravniOsnovi where isnull(Aktivan, 1) = 1";
                ucitajListu(fs, listPravniOsnov, sqlKomanda);
                //pisiListu(fs, listPravniOsnov, "PravniOsnov");

                sqlKomanda = "SELECT Id FROM dbo.UcestalostiIsplata where isnull(Aktivan, 1) = 1";
                ucitajListu(fs, listUcestalostIsplate, sqlKomanda);
                //pisiListu(fs, listPravniOsnov, "UcestalostIsplate");

                sqlKomanda = "SELECT Id FROM dbo.VrsteAkta";
                ucitajListu(fs, listVrstaAkta, sqlKomanda);
                //pisiListu(fs, listVrstaAkta, "VrstaAkta");

                sqlKomanda = "SELECT Id FROM dbo.Srodstva";
                ucitajListu(fs, listSrodstvo, sqlKomanda);
                //pisiListu(fs, listSrodstvo, "Srodstvo");

                sqlKomanda = "SELECT Id FROM dbo.drzave where isnull(Aktivan, 1) = 1";
                ucitajListu(fs, listDrzavljanstvo, sqlKomanda);
                //pisiListu(fs, listDrzavljanstvo, "Drzavljanstvo");

                sqlKomanda = "SELECT convert(int, statistika) Id FROM dbo.dpzs where isnull(Aktivan, 1) = 1 and dpzTipId = 4 and statistika is not null";
                ucitajListu(fs, listOpcina, sqlKomanda);
                //pisiListu(fs, listOpcina, "Opcina");

                //sqlKomanda = "SELECT Id FROM dbo.Prava where isnull(Aktivan, 1) = 1";
                //ucitajListu(fs, listPrava, sqlKomanda);


                sqlKomanda = "SELECT Id FROM [dbo].[GrupePrava] where isnull(Aktivan, 1) = 1";
                ucitajListu(fs, listVrstaNaknade, sqlKomanda);
                //pisiListu(fs, listPrava, "Prava");

               // getUserId(fs, parametri, "admin");
            }
            catch (Exception e)
            {
                greskaTekst = "podaciInit - Greska - " + sqlKomanda + " " + e.Message + Environment.NewLine;
                upisiGresku(listGresaka, greskaTekst);
                fs.WriteLine(greskaTekst);
                greska = true;
                fs.Flush();
            }
        }

        public void podaciInit1(StreamWriter fs, ParametriXml parametri)
        {
            brojGresaka = 0;
            numberFormat = (System.Globalization.NumberFormatInfo)currentCulture.NumberFormat.Clone();
            numberFormat.NumberDecimalSeparator = ".";

            DateTime localDate = DateTime.Now;
            godinaStr = localDate.ToString("yyyy");
            mjesecStr = localDate.ToString("MM");
            danStr = localDate.ToString("dd");



            // listeErrClear();


            createConnection(fs, parametri);
            try
            {
                listPrava.Clear();

                sqlKomanda = "SELECT p.Id FROM dbo.Prava p join dbo.PravoPodGrupe y on p.PravoPodGrupaId = y.id " +
                                " join dbo.GrupePrava g on y.GrupaPravaId = g.id " +
                                " where isnull(p.Aktivan, 1) = 1 and g.id = " + pravniOsnovId.ToString();
                ucitajListu(fs, listPrava, sqlKomanda);



            }
            catch (Exception e)
            {
                greskaTekst = "podaciInit - Greska - " + sqlKomanda + " " + e.Message + Environment.NewLine;
                upisiGresku(listGresaka, greskaTekst);
                fs.WriteLine(greskaTekst);
                greska = true;
                fs.Flush();
            }
        }
        bool postojiLice(String jmb)
        {
            bool korisnikPostoji = false;
            string komanda = "select id from dbo.lica where jmbg = @jmb";
            liceId = 0;
            try
            {
                connection.Open();
                SqlCommand command = new SqlCommand(komanda, connection);
                command.CommandType = System.Data.CommandType.Text;
                command.Parameters.AddWithValue("@jmb", jmb);

                object id = command.ExecuteScalar();
                if (id == null)
                {
                    liceId = (int?)null;
                }
                else
                {
                    liceId = Convert.ToInt32(id);
                    korisnikPostoji = true;
                }
                connection.Close();
            }
            catch
            {
           
                connection.Close();
               
            }
            return korisnikPostoji;
        }

        public int? dajOpcinu(String statistika)
        {
            int? idOpcine = 0;
            string komanda = "SELECT Id FROM dbo.dpzs where isnull(Aktivan, 1) = 1 and dpzTipId = 4 and statistika  = @statistika";
            try
            {
                connection.Open();
                SqlCommand command = new SqlCommand(komanda, connection);
                command.CommandType = System.Data.CommandType.Text;
                command.Parameters.AddWithValue("@statistika", statistika);

                object id = command.ExecuteScalar();
                if (id == null)
                {
                    idOpcine = (int?)null;
                }
                else
                {
                    idOpcine = Convert.ToInt32(id);
                    
                }
                connection.Close();
            }
            catch
            {

                connection.Close();

            }
            return idOpcine;
        
    }
      

        public Boolean kontrolaJbmg(string jmb)
        {

            try
            {
                jmbDan = Convert.ToInt32(jmb.Substring(0, 2));
                jmbMjesec = Convert.ToInt32(jmb.Substring(2, 2));
                jmbGodina = Convert.ToInt32(jmb.Substring(4, 3));
                if (jmbGodina < 100)
                {
                    jmbGodina = jmbGodina + 2000;
                }
                else
                {
                    jmbGodina = jmbGodina + 1000;
                }
                if (jmbGodina % 400 == 0)
                {
                    isPrestupna = true;
                }
                else if (jmbGodina % 100 == 0)
                {
                    isPrestupna = false;
                }
                else if (jmbGodina % 4 == 0)
                {
                    isPrestupna = true;
                }
                else
                {
                    isPrestupna = false;
                }
                if (jmbMjesec == 1 || jmbMjesec == 3 || jmbMjesec == 5 || jmbMjesec == 7 || jmbMjesec == 8 || jmbMjesec == 10 || jmbMjesec == 12)
                {
                    maxDana = 31;
                }
                else if (jmbMjesec == 4 || jmbMjesec == 6 || jmbMjesec == 9 || jmbMjesec == 11)
                {
                    maxDana = 30;
                }
                else if (jmbMjesec == 2 && isPrestupna == true)
                {
                    maxDana = 29;
                }
                else if (jmbMjesec == 2 && isPrestupna == false)
                {
                    maxDana = 28;
                }
                else
                {
                    greska = true;
                    upisiGresku(listGresaka, "obradiKorisnika - Greska - pogresan mjesec JMB " + jmb);
                }
                if (jmbGodina <= 1900 || jmbGodina > maxGodina)
                {
                    greska = true;
                    upisiGresku(listGresaka, "obradiKorisnika - Greska - pogresna godina JMB " + jmb);
                }
                if (jmbDan < 1 || jmbDan > maxDana)
                {
                    greska = true;
                    upisiGresku(listGresaka, "obradiKorisnika - Greska - pogresan dan JMB " + jmb);
                }
            }
            catch (Exception e)
            {
                greska = true;
                upisiGresku(listGresaka, "obradiKorisnika - Greska - greska u konverziji JMB " + jmb);
            }

            return greska;
        }


        public Boolean testFirmaIzNaziva(StreamWriter fs, ParametriXml par, string fileName)
        {
            bool isOk = false;
            string polje = "";
            institucija = new Institucija();
            //           JR4200000000001_20190621123822

            justName = Path.GetFileNameWithoutExtension(fileName);
            jibName = justName.Substring(2).Substring(0, 13);
            jibDadoteka = jibName;
            institucija.jib = jibName;
            //            string komanda = "select id from dbo.institucije where jib = @jib";
            string komanda = "select id, naziv, isnull(EMail, '') eMail from dbo.institucije where jib = @jib";

            try
            {
                connection.Open();
                SqlCommand command = new SqlCommand(komanda, connection);
                command.CommandType = System.Data.CommandType.Text;
                command.Parameters.AddWithValue("@jib", jibName);

                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    polje = "id";
                    if (!reader.IsDBNull(reader.GetOrdinal(polje)))
                    {
                        firmaId = reader.GetInt32(reader.GetOrdinal(polje));
                        isOk = true;
                    }
                    else
                    {
                        firmaId = null;
                    }

                    polje = "naziv";
                    if (!reader.IsDBNull(reader.GetOrdinal(polje)))
                    {
                        institucija.NazivInstitucije = reader.GetString(reader.GetOrdinal(polje));
                    }
                    else
                    {
                        institucija.NazivInstitucije = "Nepoznata institucija";
                        isOk = false;
                    }

                    //polje = "eMail";
                    //if (!reader.IsDBNull(reader.GetOrdinal(polje)))
                    //{
                    //    institucija.eMailInstitucije = reader.GetString(reader.GetOrdinal(polje));
                    //    if (institucija.eMailInstitucije != "")
                    //    {
                    //        //                            listeMail.Add(institucija.eMailInstitucije);
                    //        if (!eMail.Equals(institucija.eMailInstitucije))
                    //            addEMail(listeMail, institucija.eMailInstitucije, " - institucija");
                    //    }
                    //}
                }
            }
            catch (Exception e)
            {
                greskaTekst = "testFirmaIzNaziva - Greska " + e.Message;
                upisiGresku(listGresaka, greskaTekst);
                fs.WriteLine(greskaTekst);
                greska = true;
                fs.Flush();
                isOk = false;
            }
            connection.Close();
            return isOk;
        }
        public void addEMail(List<string> listeMail, string eMail, string izvor)
        {

            if (ru.IsValidEmail(eMail) == false)
            {
                upisiGresku(listUpozorenja, "Mail " + eMail + " nije korektan" + izvor);
                return;
            }

            foreach (string s in listeMail)
            {
                if (s == eMail)
                {
                    return;
                }
            }

            listeMail.Add(eMail);
            return;
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

        public object[,] otvoriExcelDat(String nazivDatoteke, String nazivXmlDatoteke)
        {
            //Microsoft.Office.Interop.Excel.Application 
                xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkbook = xlApp.Workbooks.Open(nazivDatoteke);
            xlWorksheet = xlWorkbook.Sheets[1];
            Range xlRange = xlWorksheet.UsedRange;

            int brojkolona = xlRange.Columns.Count;
            int brojredova = xlRange.Rows.Count;
            k = 0;
            ukupanIznosStavki = 0.00;
            ukupanBrojStavki = 0;
            // redovi sa isplatama
            objectArray = (object[,])xlRange.Value[XlRangeValueDataType.xlRangeValueDefault];
            return objectArray;
        }

        public void zatvoriExcel(Workbook wb)
        {
            wb.Close();            
        }


        public int provjeraExcelDatoteke(String nazivDatoteke, StreamWriter fs, StreamWriter fs1, ParametriXml par, String nazivXmlDatoteke)
        {
            try
            {
                brojGresaka = 0;
                int ukupanBroj;
            // niz sa redovima iz excela
            //objectArray = otvoriExcelDat(nazivDatoteke, nazivXmlDatoteke);
            //Microsoft.Office.Interop.Excel.Application
                xlApp = new Microsoft.Office.Interop.Excel.Application();
            
                xlWorkbook = xlApp.Workbooks.Open(nazivDatoteke);
                xlWorksheet = xlWorkbook.Sheets[1];
                Range xlRange = xlWorksheet.UsedRange;
                int brojkolona = xlRange.Columns.Count;
                int brojredova = xlRange.Rows.Count;
                k = 0;
                ukupanIznosStavki = 0.00;
                ukupanBrojStavki = 0;
                // redovi sa isplatama
                objectArray = (object[,])xlRange.Value[XlRangeValueDataType.xlRangeValueDefault];

                // zaglavlje - citanje 
                jib = (objectArray[2, 1] == null) ? string.Empty : objectArray[2, 1].ToString();
                eMail = (objectArray[2, 2] == null) ? string.Empty : objectArray[2, 2].ToString();               
                addEMail(listeMail, eMail, " - Excel");
                //provjera
                grupaPravaId = (objectArray[2, 3] == null) ? string.Empty : objectArray[2, 3].ToString();
                //provjera
                pravniOsnovId = (objectArray[2, 4] == null) ? string.Empty : objectArray[2, 4].ToString();
                podaciInit1(fs, par);
                //provjera
                ucestalostIsplataId = (objectArray[2, 5] == null) ? string.Empty : objectArray[2, 5].ToString();
                try
                {
                    datumIsplate1 = Convert.ToDateTime((objectArray[2, 6] == null) ? string.Empty : objectArray[2, 6].ToString());
                    datumIsplate = datumIsplate1.ToString("yyyy-MM-dd").ToString();
                }
                catch (Exception ee)
                {
                    datumIsplate = "2000-01-01";//.ToString("yyyy-MM-dd").ToString();
                }
                brojIsplata = (objectArray[2, 7] == null) ? string.Empty : objectArray[2, 7].ToString().Replace(",", "."); 
                ukupanIznos = (objectArray[2, 8] == null) ? string.Empty : objectArray[2, 8].ToString();
                nazivInstitucije = (objectArray[2, 9] == null) ? string.Empty : objectArray[2, 9].ToString();
                // upis na pocetak obavijesti
                fs1.WriteLine("Naziv institucije: " + nazivInstitucije);
                fs1.Flush();
                try
                {
                    fs1.WriteLine("Kategorija prava: " + getNazivKategorijaPrava(Int32.Parse(pravniOsnovId)));
                }
                catch(Exception e)
                {                
                    KillExcel(xlApp);
                    System.Threading.Thread.Sleep(500);
                    moveFile(fs, par, nazivXmlDatoteke, true);
                    string poruka = "Pogresna kategorija prava: " + nazivXmlDatoteke;

                    listaGresaka = new String[100];
                    listaGresaka[brojGresaka] = poruka;
                    brojGresaka++;
                    // upisati u neki fajl greske
                    fs1.WriteLine("Greska " + brojGresaka.ToString() + "-" + poruka);
                    fs1.Flush();

                    String imeGreske = ((FileStream)(fs1.BaseStream)).Name;
                    fs1.Close();
                    fs1.Dispose();
                   // posaljiMail(fs, par, imeGreske, "Problem sa kategorijom prava u zaglavlju "); //"GRESKA TEST"); //listaGresaka.ToString());
                    return brojGresaka;
                }
                // iznosi
                //fs1.WriteLine("Broj isplata: " + brojIsplata.ToString() + " Iznos: " + ukupanIznos.ToString());
                //fs1.Flush();
                // provjera podataka iz zaglavlja
                //kontrola da li postoje clanovi
                try
                {
                    if (objectArray[3, 20] != null)
                    {
                        c1 = true;
                    }
                } catch (Exception cc)
                { 
                    c1 = false;
                }
                    Boolean isprav =  kontrolaZaglavlja(fs, fs1, par, nazivXmlDatoteke);
                ispravna = true;
                try
                {
                    ukupanBroj = Int32.Parse(brojIsplata); // objectArray.GetUpperBound(0) - 3;
                }
                catch(Exception ee)
                {
                    KillExcel(xlApp);
                    System.Threading.Thread.Sleep(500);
                    moveFile(fs, par, nazivXmlDatoteke, true);
                    string poruka = "Pogresan broj isplata : " + nazivXmlDatoteke;

                    listaGresaka = new String[100];
                    listaGresaka[brojGresaka] = poruka;
                    brojGresaka++;
                    // upisati u neki fajl greske
                    fs1.WriteLine("Greska " + brojGresaka.ToString() + "-" + poruka);
                    fs1.Flush();

                    String imeGreske = ((FileStream)(fs1.BaseStream)).Name;
                    fs1.Close();
                    fs1.Dispose();
                  //  posaljiMail(fs, par, imeGreske, "Problem sa kategorijom prava u zaglavlju "); //"GRESKA TEST"); //listaGresaka.ToString());
                    return brojGresaka;
                    //ukupanBroj = 0;
                    //brojGresaka++;
                }
                if (isprav == true)
                {
                    // procedura za obradu jednog dokumenta
                    for (row = 4; row < 4 + ukupanBroj; row++)//row <= objectArray.GetUpperBound(0); row++)
                    {
                        try
                        {
                            obradaJedneIsplate(fs1, objectArray, row, xlWorkbook);

                            kontrolaSvihPodatakaIzIsplata(fs, fs1, par, row, ukupanBroj, nazivXmlDatoteke);
                            // xlWorkbook.Close();
                        }
                        catch (Exception e)
                        {


                            fs1.WriteLine("Greska - procedura za obradu jednog dokumenta " + e.Message);
                            fs1.Flush();
                        }
                    }
                }
            }
            catch (Exception ex) {
               
                fs1.WriteLine("Greska - provjeraExcelDatoteke " + ex.Message);
                fs1.Flush();
                brojGresaka = 1;
            }
            finally
            {
              //  KillExcel(xlApp);
                //System.Threading.Thread.Sleep(100);
            }



            return brojGresaka;
        }

        public string getNazivKategorijaPrava(int id)
        {
            sqlKomanda = "SELECT naziv FROM dbo.GrupePrava where id = @Id";
            string naziv = "";
            try
            {
                connection.Open();
                SqlCommand command = new SqlCommand(sqlKomanda, connection);

                command.Parameters.AddWithValue("@Id", id);

                object nazivO = command.ExecuteScalar();
                naziv = (string)nazivO;
                connection.Close();
            }
            catch
            {
                naziv = "?";
            }
            return naziv;
        }

        [DllImport("User32.dll")]
        public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int ProcessId);
        private static void KillExcel(Microsoft.Office.Interop.Excel.Application theApp)
        {
            int id = 0;
            IntPtr intptr = new IntPtr(theApp.Hwnd);
            System.Diagnostics.Process p = null;
            try
            {
                GetWindowThreadProcessId(intptr, out id);
                p = System.Diagnostics.Process.GetProcessById(id);
                if (p != null)
                {
                    p.Kill();
                    p.Dispose();
                }
            }
            catch (Exception ex)
            {
              //  System.Windows.Forms.MessageBox.Show("KillExcel:" + ex.Message);
            }
        }

        public void obradaJedneIsplate(StreamWriter fs, object[,] objectArray1, int row, Workbook xlWorkbook1)
        {           
            try
            {
                //sve su stringovi, kasnije u int konvertovati sto treba
                redniBroj = (objectArray1[row, 1] == null) ? string.Empty : objectArray1[row, 1].ToString();
                jmb = (objectArray1[row, 2] == null) ? string.Empty : objectArray1[row, 2].ToString().Trim();
                prezime = (objectArray1[row, 3] == null) ? string.Empty : objectArray1[row, 3].ToString();
                ime = (objectArray1[row, 4] == null) ? string.Empty : objectArray1[row, 4].ToString();
                roditelj = (objectArray1[row, 5] == null) ? string.Empty : objectArray1[row, 5].ToString();
                mjestoRodjenja = (objectArray1[row, 6] == null) ? string.Empty : objectArray1[row, 6].ToString();

                //provjera
                drzavljanstvoId = (objectArray1[row, 7] == null) ? string.Empty : objectArray1[row, 7].ToString();
                opcinaId = (objectArray1[row, 8] == null) ? string.Empty : objectArray1[row, 8].ToString();
                posta = (objectArray1[row, 9] == null) ? string.Empty : objectArray1[row, 9].ToString();
                mjesto = (objectArray1[row, 10] == null) ? string.Empty : objectArray1[row, 10].ToString();
                adresa = (objectArray1[row, 11] == null) ? string.Empty : objectArray1[row, 11].ToString();
                //provjera
                pravoId = (objectArray1[row, 12] == null) ? string.Empty : objectArray1[row, 12].ToString();
                iznos = (objectArray1[row, 13] == null) ? string.Empty : objectArray1[row, 13].ToString();//.Replace(",", ".");
                //provjera
                vrstaAktaId = (objectArray1[row, 14] == null) ? string.Empty : objectArray1[row, 14].ToString();
                protokol = (objectArray1[row, 15] == null) ? string.Empty : objectArray1[row, 15].ToString();
                try
                {
                    DateTime datumAkta1 = Convert.ToDateTime((objectArray1[row, 16] == null) ? string.Empty : objectArray1[row, 16].ToString());
                    datumAkta = datumAkta1.ToString("yyyy-MM-dd").ToString();
                }
                catch (Exception ee)
                {
                    datumAkta = "2000-01-01";//.ToString("yyyy-MM-dd").ToString();
                }

                odgovornoLice = (objectArray1[row, 17] == null) ? string.Empty : objectArray1[row, 17].ToString();
                unosLice = (objectArray1[row, 18] == null) ? string.Empty : objectArray1[row, 18].ToString();

                //domacinstva
                clanPrezime1 = (objectArray1[row, 19] == null) ? string.Empty : objectArray1[row, 19].ToString();
                clanIme1 = (objectArray1[row, 20] == null) ? string.Empty : objectArray1[row, 20].ToString();
                clanSrod1 = (objectArray1[row, 21] == null) ? string.Empty : objectArray1[row, 21].ToString();
                clanGodiste1 = (objectArray1[row, 22] == null) ? string.Empty : objectArray1[row, 22].ToString();
                vclanLK1 = (objectArray1[row, 23] == null) ? string.Empty : objectArray1[row, 23].ToString();

                clanPrezime2 = (objectArray1[row, 24] == null) ? string.Empty : objectArray1[row, 24].ToString();
                clanIme2 = (objectArray1[row, 25] == null) ? string.Empty : objectArray1[row, 25].ToString();
                clanSrod2 = (objectArray1[row, 26] == null) ? string.Empty : objectArray1[row, 26].ToString();
                clanGodiste2 = (objectArray1[row, 27] == null) ? string.Empty : objectArray1[row, 27].ToString();
                vclanLK2 = (objectArray1[row, 28] == null) ? string.Empty : objectArray1[row, 28].ToString();

                clanPrezime3 = (objectArray1[row, 29] == null) ? string.Empty : objectArray1[row, 29].ToString();
                clanIme3 = (objectArray1[row, 30] == null) ? string.Empty : objectArray1[row, 30].ToString();
                clanSrod3 = (objectArray1[row, 31] == null) ? string.Empty : objectArray1[row, 31].ToString();
                clanGodiste3 = (objectArray1[row, 32] == null) ? string.Empty : objectArray1[row, 32].ToString();
                vclanLK3 = (objectArray1[row, 33] == null) ? string.Empty : objectArray1[row, 33].ToString();

                clanPrezime4 = (objectArray1[row, 34] == null) ? string.Empty : objectArray1[row, 34].ToString();
                clanIme4 = (objectArray1[row, 35] == null) ? string.Empty : objectArray1[row, 35].ToString();
                clanSrod4 = (objectArray1[row, 36] == null) ? string.Empty : objectArray1[row, 36].ToString();
                clanGodiste4 = (objectArray1[row, 37] == null) ? string.Empty : objectArray1[row, 37].ToString();
                vclanLK4 = (objectArray1[row, 38] == null) ? string.Empty : objectArray1[row, 38].ToString();

                clanPrezime5 = (objectArray1[row, 39] == null) ? string.Empty : objectArray1[row, 39].ToString();
                clanIme5 = (objectArray1[row, 40] == null) ? string.Empty : objectArray1[row, 40].ToString();
                clanSrod5 = (objectArray1[row, 41] == null) ? string.Empty : objectArray1[row, 41].ToString();
                clanGodiste5 = (objectArray1[row, 42] == null) ? string.Empty : objectArray1[row, 42].ToString();
                vclanLK5 = (objectArray1[row, 43] == null) ? string.Empty : objectArray1[row, 43].ToString();

                try
                {
                    ukupanIznosStavki = ukupanIznosStavki + Double.Parse(iznos);
                }
                catch (Exception eee)
                {
                    ukupanIznosStavki = ukupanIznosStavki;
                }
                ukupanBrojStavki = ukupanBrojStavki + 1;

                // kontrola
                //   kontrolaSvihPodataka(fs, fs1, par, ime);
                //  xlWorkbook1.Close();
            }
            catch (Exception e)
            {

               fs.WriteLine("Greska obradaJedneIsplate " + e.Message);
               fs.Flush();

            }
        }

        public void upuniListeKontrolnihVrijednosti(StreamWriter fs, ParametriXml parametri)
        {
            createConnection(fs, parametri);
            try
            {
                sqlKomanda = "SELECT Id FROM dbo.PravniOsnovi where isnull(Aktivan, 1) = 1";
                ucitajListu(fs, listPravniOsnov, sqlKomanda);
                //pisiListu(fs, listPravniOsnov, "PravniOsnov");

                sqlKomanda = "SELECT Id FROM dbo.UcestalostiIsplata where isnull(Aktivan, 1) = 1";
                ucitajListu(fs, listUcestalostIsplate, sqlKomanda);
                //pisiListu(fs, listPravniOsnov, "UcestalostIsplate");

                sqlKomanda = "SELECT Id FROM dbo.VrsteAkta";
                ucitajListu(fs, listVrstaAkta, sqlKomanda);
                //pisiListu(fs, listVrstaAkta, "VrstaAkta");

                sqlKomanda = "SELECT Id FROM dbo.Srodstva";
                ucitajListu(fs, listSrodstvo, sqlKomanda);
                //pisiListu(fs, listSrodstvo, "Srodstvo");

                sqlKomanda = "SELECT Id FROM dbo.drzave where isnull(Aktivan, 1) = 1";
                ucitajListu(fs, listDrzavljanstvo, sqlKomanda);
                //pisiListu(fs, listDrzavljanstvo, "Drzavljanstvo");

                sqlKomanda = "SELECT convert(int, statistika) Id FROM dbo.dpzs where isnull(Aktivan, 1) = 1 and dpzTipId = 4 and statistika is not null";
                ucitajListu(fs, listOpcina, sqlKomanda);
                //pisiListu(fs, listOpcina, "Opcina");

                sqlKomanda = "SELECT Id FROM dbo.Prava where isnull(Aktivan, 1) = 1";
                ucitajListu(fs, listPrava, sqlKomanda);
                //pisiListu(fs, listPrava, "Prava");

                sqlKomanda = "SELECT Id FROM [dbo].[GrupePrava] where isnull(Aktivan, 1) = 1";
                ucitajListu(fs, listVrstaNaknade, sqlKomanda);
                //pisiListu(fs, listPrava, "Prava");
            }
            catch (Exception e)
            {
                greskaTekst = "podaciInit - Greska - " + sqlKomanda + " " + e.Message + Environment.NewLine;
                upisiGresku(listGresaka, greskaTekst);
                fs.WriteLine(greskaTekst);
                greska = true;
                fs.Flush();
            }
        }



        public Boolean provjeraMaticnog(StreamWriter fs1, String jmb)
        {
            greska = false;
            if (jmb.Trim().Length != 13)
            {
                greska = true;
                upisiGresku(listGresaka, "obradiKorisnika - Greska - pogresna duzina JMB " + jmb);
                fs1.WriteLine("Pogresna duzina JMB" + jmb);
                fs1.Flush();
            }
            else
            {
                try
                {
                    jmbDan = Convert.ToInt32(jmb.Substring(0, 2));
                    jmbMjesec = Convert.ToInt32(jmb.Substring(2, 2));
                    jmbGodina = Convert.ToInt32(jmb.Substring(4, 3));
                    if (jmbGodina < 100)
                    {
                        jmbGodina = jmbGodina + 2000;
                    }
                    else
                    {
                        jmbGodina = jmbGodina + 1000;
                    }
                    if (jmbGodina % 400 == 0)
                    {
                        isPrestupna = true;
                    }
                    else if (jmbGodina % 100 == 0)
                    {
                        isPrestupna = false;
                    }
                    else if (jmbGodina % 4 == 0)
                    {
                        isPrestupna = true;
                    }
                    else
                    {
                        isPrestupna = false;
                    }
                    if (jmbMjesec == 1 || jmbMjesec == 3 || jmbMjesec == 5 || jmbMjesec == 7 || jmbMjesec == 8 || jmbMjesec == 10 || jmbMjesec == 12)
                    {
                        maxDana = 31;
                    }
                    else if (jmbMjesec == 4 || jmbMjesec == 6 || jmbMjesec == 9 || jmbMjesec == 11)
                    {
                        maxDana = 30;
                    }
                    else if (jmbMjesec == 2 && isPrestupna == true)
                    {
                        maxDana = 29;
                    }
                    else if (jmbMjesec == 2 && isPrestupna == false)
                    {
                        maxDana = 28;
                    }
                    else
                    {
                        greska = true;
                        upisiGresku(listGresaka, "obradiKorisnika - Greska - pogresan mjesec JMB " + jmb);
                        fs1.WriteLine("Greska pogresan mjesec JMB" + jmb);
                        fs1.Flush();
                    }
                    if (jmbGodina <= 1900 || jmbGodina > maxGodina)
                    {
                        greska = true;
                        upisiGresku(listGresaka, "obradiKorisnika - Greska - pogresna godina JMB " + jmb);
                        fs1.WriteLine("Greska - pogresna godina JMB" + jmb);
                        fs1.Flush();
                    }
                    if (jmbDan < 1 || jmbDan > maxDana)
                    {
                        greska = true;
                        upisiGresku(listGresaka, "obradiKorisnika - Greska - pogresan dan JMB " + jmb);
                        fs1.WriteLine("Greska - pogresan dan JMB" + jmb);
                        fs1.Flush();
                    }
                }
                catch (Exception e)
                {
                    greska = true;
                    upisiGresku(listGresaka, "obradiKorisnika - Greska - greska u konverziji JMB " + jmb);
                }
            }
            return greska;
        }

        void ucitajListu(StreamWriter fs, List<int> l, string komanda)
        {
            var command = new SqlCommand(komanda, this.connection);

            try
            {
                connection.Open();

                SqlDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    int Id = reader.GetInt32(reader.GetOrdinal("Id"));
                    l.Add(Id);
                }

                connection.Close();
            }
            catch (Exception e)
            {              
                fs.WriteLine(e.Message);
                fs.Flush();
                throw e;
            }

        }
        // konekcija
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

        public Boolean kontrolaZaglavlja(StreamWriter fs, StreamWriter fs1, ParametriXml par, String nazivXmlDatoteke)
        {
            Boolean ispravan = true; 
            userId = (int?)null;
            fs1.WriteLine("PRENOS dokumenta " + nazivXmlDatoteke);
            fs1.WriteLine("--------------------------------------");
            fs1.Flush();
           
            listaGresaka = new String[100];

            if (userId == null)
            {
                if (getUserId(fs, par, "admin") == false)
                {
                    string poruka = "Nema admin korisnika: " + nazivXmlDatoteke;


                    listaGresaka[brojGresaka] = poruka;
                    brojGresaka++;
                }
            }
            if (testFirmaIzNaziva(fs, par, nazivXmlDatoteke) == false)
            {
                ispravan = false;
                KillExcel(xlApp);
                System.Threading.Thread.Sleep(500);
                moveFile(fs, par, nazivXmlDatoteke, true);
                string poruka = "Nema firme sa JIB-om iz naziva datoteke: " + nazivXmlDatoteke + " (" + jib + ")";


                listaGresaka[brojGresaka] = poruka;
                brojGresaka++;
                // upisati u neki fajl greske
                fs1.WriteLine("Greska " + brojGresaka.ToString() + "-" + poruka);
                fs1.Flush();

                String imeGreske = ((FileStream)(fs1.BaseStream)).Name;
                fs1.Close();
                fs1.Dispose();
               // posaljiMail(fs, par, imeGreske, "JIB problem"); //"GRESKA TEST"); //listaGresaka.ToString());
                return ispravan;                
            }
            else
            {

            }

            if (datumIsplate.Equals("2000-01-01")) 
            {
                ispravan = false;
                KillExcel(xlApp);
                System.Threading.Thread.Sleep(500);
                moveFile(fs, par, nazivXmlDatoteke, true);
                string poruka = "Pogrešan datum isplate: " + nazivXmlDatoteke ;


                listaGresaka[brojGresaka] = poruka;
                brojGresaka++;
                // upisati u neki fajl greske
                fs1.WriteLine("Greska " + brojGresaka.ToString() + "-" + poruka);
                fs1.Flush();

                String imeGreske = ((FileStream)(fs1.BaseStream)).Name;
                fs1.Close();
                fs1.Dispose();
              //   posaljiMail(fs, par, imeGreske, "Pogrešan datum isplate"); //"GRESKA TEST"); //listaGresaka.ToString());
                return ispravan;
            }
            else
            {

            }
            // pravni osnov
            try
            {
                if (!listPravniOsnov.Contains(Int32.Parse(grupaPravaId)))
                {
                    string poruka = "Ne postoji pravni osnov: " + grupaPravaId;


                    listaGresaka[brojGresaka] = poruka;
                    brojGresaka++;
                    // upisati u neki fajl greske
                    fs1.WriteLine("Greska " + brojGresaka.ToString() + "-" + poruka);
                    fs1.Flush();
                }
            } 
            catch (Exception ee)
            {
                ispravan = false;
                KillExcel(xlApp);
                System.Threading.Thread.Sleep(500);
                moveFile(fs, par, nazivXmlDatoteke, true);
                string poruka = "Problem sa pravnim osnovom u zaglavlju: " + nazivXmlDatoteke + "- " + grupaPravaId;


                listaGresaka[brojGresaka] = poruka;
                brojGresaka++;
                // upisati u neki fajl greske
                fs1.WriteLine("Greska " + brojGresaka.ToString() + "-" + poruka);
                fs1.Flush();

                String imeGreske = ((FileStream)(fs1.BaseStream)).Name;
                fs1.Close();
                fs1.Dispose();
               // posaljiMail(fs, par, imeGreske, "Problem sa pravnim osnovom u zaglavlju"); //"GRESKA TEST"); //listaGresaka.ToString());
                return ispravan;
            }


            //kategorija naknade
            try
            {
                if (!listVrstaNaknade.Contains(Int32.Parse(pravniOsnovId)))
                {
                    string poruka = "Ne postoji kategorija naknade: " + pravniOsnovId;


                    listaGresaka[brojGresaka] = poruka;
                    brojGresaka++;
                    // upisati u neki fajl greske
                    fs1.WriteLine("Greska " + brojGresaka.ToString() + "-" + poruka);
                    fs1.Flush();
                }
            } catch (Exception ee)
            {
                ispravan = false;
                KillExcel(xlApp);
                System.Threading.Thread.Sleep(500);
                moveFile(fs, par, nazivXmlDatoteke, true);
                string poruka = "Problem sa kategorijom naknade u zaglavlju: " + nazivXmlDatoteke + "- " + pravniOsnovId;


                listaGresaka[brojGresaka] = poruka;
                brojGresaka++;
                // upisati u neki fajl greske
                fs1.WriteLine("Greska " + brojGresaka.ToString() + "-" + poruka);
                fs1.Flush();

                String imeGreske = ((FileStream)(fs1.BaseStream)).Name;
                fs1.Close();
                fs1.Dispose();
              //  posaljiMail(fs, par, imeGreske, "Problem sa kategorijom naknade u zaglavlju"); //"GRESKA TEST"); //listaGresaka.ToString());
                return ispravan;
            }
                        
                if (c1 == true)
                {

                }
                else
                {
                ispravan = false;
                KillExcel(xlApp);
                System.Threading.Thread.Sleep(500);
                moveFile(fs, par, nazivXmlDatoteke, true);
                string poruka = "Problem sa zaglavljem dokumenta, nedostaju kolone: " + nazivXmlDatoteke;


                listaGresaka[brojGresaka] = poruka;
                brojGresaka++;
                // upisati u neki fajl greske
                fs1.WriteLine("Greska " + brojGresaka.ToString() + "-" + poruka);
                fs1.Flush();

                String imeGreske = ((FileStream)(fs1.BaseStream)).Name;
                fs1.Close();
                fs1.Dispose();
              //  posaljiMail(fs, par, imeGreske, "Problem sa zaglavljem dokumenta, nedostaju kolone"); //"GRESKA TEST"); //listaGresaka.ToString());
                return ispravan;
            }
            try
            {
                // ucestalost
                if (!listUcestalostIsplate.Contains(Int32.Parse(ucestalostIsplataId)))
                {
                    string poruka = "Ne postoji ucestalost isplate: " + ucestalostIsplataId;


                    listaGresaka[brojGresaka] = poruka;
                    brojGresaka++;
                    // upisati u neki fajl greske
                    fs1.WriteLine("Greska " + brojGresaka.ToString() + "-" + poruka);
                    fs1.Flush();
                }
            }
            catch (Exception ee)
            {

            }
            return ispravan;
        }


        public void kontrolaSvihPodatakaIzIsplata(StreamWriter fs, StreamWriter fs1, ParametriXml par, int red, int brojisplata, String nazivDat)
        {
            if (ispravna == true)
            {
                try
                {
                    k++;

                    if (provjeraMaticnog(fs1, jmb) == true)
                    {
                        string poruka = "Greska u jmbg " + jmb;
                        listaGresaka[brojGresaka] = poruka;
                        brojGresaka++;
                        fs1.WriteLine("Greska " + brojGresaka.ToString() + "-" + poruka + " u redu " + row);
                        fs1.Flush();
                    }
                    else
                    {

                    }
                    //drzavljanstvo
                    try
                    {
                        if (!listDrzavljanstvo.Contains(Int32.Parse(drzavljanstvoId)))
                        {
                            string poruka = "Ne postoji drzava: " + drzavljanstvoId;


                            listaGresaka[brojGresaka] = poruka;
                            brojGresaka++;
                            // upisati u neki fajl greske
                            fs1.WriteLine("Greska " + brojGresaka.ToString() + "-" + poruka + " u redu " + row);
                            fs1.Flush();
                        }
                    }
                    catch (Exception e)
                    {
                        listaGresaka[brojGresaka] = "Ne postoji drzava: " + drzavljanstvoId;
                        brojGresaka++;
                        fs1.WriteLine("Greska " + brojGresaka.ToString() + "-" + "Ne postoji drzava: " + drzavljanstvoId + " u redu " + row);
                        fs1.Flush();
                    }

                    // opcina
                    try
                    {
                        if (!listOpcina.Contains(Int32.Parse(opcinaId)))
                        {
                            string poruka = "Ne postoji opcina: " + opcinaId;


                            listaGresaka[brojGresaka] = poruka;
                            brojGresaka++;
                            // upisati u neki fajl greske
                            fs1.WriteLine("Greska " + brojGresaka.ToString() + "-" + poruka + " u redu " + row);
                            fs1.Flush();
                        }
                    }
                    catch (Exception e)
                    {
                        listaGresaka[brojGresaka] = "Ne postoji opcina: " + opcinaId;
                        brojGresaka++;
                        fs1.WriteLine("Greska " + brojGresaka.ToString() + "-" + "Ne postoji opcina: " + opcinaId + " u redu " + row);
                        fs1.Flush();
                    }
                    //pravo
                    try
                    {
                        if (!listPrava.Contains(Int32.Parse(pravoId)))
                        {
                            string poruka = "Ne postoji pravo: " + pravoId;


                            listaGresaka[brojGresaka] = poruka;
                            brojGresaka++;
                            // upisati u neki fajl greske
                            fs1.WriteLine("Greska " + brojGresaka.ToString() + "-" + poruka + " u redu " + row);
                            fs1.Flush();
                        }
                    }
                    catch (Exception e)
                    {
                        string poruka = "Ne postoji pravo: " + pravoId;


                        listaGresaka[brojGresaka] = poruka;
                        brojGresaka++;
                        fs1.WriteLine("Greska " + brojGresaka.ToString() + "-" + "Ne postoji pravo: " + pravoId + " u redu " + row);
                        fs1.Flush();
                    }

                    if (datumAkta.Equals("2000-01-01"))
                    {
                        string poruka = "Nije ispravan datum akta: " ;


                        listaGresaka[brojGresaka] = poruka;
                        brojGresaka++;
                        // upisati u neki fajl greske
                        fs1.WriteLine("Greska " + brojGresaka.ToString() + "-" + poruka + " u redu " + row);
                        fs1.Flush();
                    }
                        //vrsta akta
                        try
                    {
                        if (!listVrstaAkta.Contains(Int32.Parse(vrstaAktaId)))
                        {
                            string poruka = "Ne postoji vrsta akta: " + vrstaAktaId;


                            listaGresaka[brojGresaka] = poruka;
                            brojGresaka++;
                            // upisati u neki fajl greske
                            fs1.WriteLine("Greska " + brojGresaka.ToString() + "-" + poruka + " u redu " + row);
                            fs1.Flush();
                        }
                    }
                    catch (Exception e)
                    {
                        string poruka = "Ne postoji vrsta akta: " + vrstaAktaId;
                        listaGresaka[brojGresaka] = poruka;
                        brojGresaka++;
                        fs1.WriteLine("Greska " + brojGresaka.ToString() + "-" + "Ne postoji vrsta akta: " + vrstaAktaId + " u redu " + row);
                        fs1.Flush();
                    }

                    try
                    {
                        if (!clanSrod1.Equals(""))
                        {
                            try
                            {
                                //srodstvo
                                if (!listSrodstvo.Contains(Int32.Parse(clanSrod1)))
                                {
                                    string poruka = "Ne postoji srodstvo: " + clanSrod1;


                                    listaGresaka[brojGresaka] = poruka;
                                    brojGresaka++;
                                    // upisati u neki fajl greske
                                    fs1.WriteLine("Greska " + brojGresaka.ToString() + "-" + poruka + " u redu " + row);
                                    fs1.Flush();
                                }
                            }
                            catch (Exception ee)
                            {
                                string poruka = "Pogresno srodstvo: " + clanSrod1;


                                listaGresaka[brojGresaka] = poruka;
                                brojGresaka++;
                                // upisati u neki fajl greske
                                fs1.WriteLine("Greska " + brojGresaka.ToString() + "-" + poruka + " u redu " + row);
                                fs1.Flush();

                            }
                        }
                        if (!clanSrod2.Equals(""))
                        {
                            try
                            {
                                if (!listSrodstvo.Contains(Int32.Parse(clanSrod2)))
                                {                                   
                                    string poruka = "Ne postoji srodstvo: " + clanSrod2;
                                    listaGresaka[brojGresaka] = poruka;
                                    brojGresaka++;
                                    // upisati u neki fajl greske
                                    fs1.WriteLine("Greska " + brojGresaka.ToString() + "-" + poruka + " u redu " + row);
                                    fs1.Flush();
                                }
                            }
                            catch (Exception ee)
                            {
                                string poruka = "Pogresno srodstvo: " + clanSrod2;
                                listaGresaka[brojGresaka] = poruka;
                                brojGresaka++;
                                // upisati u neki fajl greske
                                fs1.WriteLine("Greska " + brojGresaka.ToString() + "-" + poruka + " u redu " + row);
                                fs1.Flush();
                            }

                }
                        if (!clanSrod3.Equals(""))
                        {
                            try
                            {
                                if (!listSrodstvo.Contains(Int32.Parse(clanSrod3)))
                                {

                                    string poruka = "Ne postoji srodstvo: " + clanSrod3;


                                    listaGresaka[brojGresaka] = poruka;
                                    brojGresaka++;
                                    // upisati u neki fajl greske
                                    fs1.WriteLine("Greska " + brojGresaka.ToString() + "-" + poruka + " u redu " + row);
                                    fs1.Flush();
                                }
                            }
                            catch (Exception ee)
                            {
                                string poruka = "Pogresno srodstvo: " + clanSrod3;


                                listaGresaka[brojGresaka] = poruka;
                                brojGresaka++;
                                // upisati u neki fajl greske
                                fs1.WriteLine("Greska " + brojGresaka.ToString() + "-" + poruka + " u redu " + row);
                                fs1.Flush();
                            }

                        }

                        if (!clanSrod4.Equals(""))
                        {
                            try
                            {
                                if (!listSrodstvo.Contains(Int32.Parse(clanSrod4)))
                                {
                                    string poruka = " Ne postoji srodstvo: " + clanSrod4;


                                    listaGresaka[brojGresaka] = poruka;
                                    brojGresaka++;
                                    // upisati u neki fajl greske
                                    fs1.WriteLine("Greska " + brojGresaka.ToString() + "-" + poruka + " u redu " + row);
                                    fs1.Flush();
                                }
                            }
                            catch (Exception ee)
                            {
                                string poruka = " Pogresno srodstvo: " + clanSrod4;


                                listaGresaka[brojGresaka] = poruka;
                                brojGresaka++;
                                // upisati u neki fajl greske
                                fs1.WriteLine("Greska " + brojGresaka.ToString() + "-" + poruka + " u redu " + row);
                                fs1.Flush();
                            }
                        }
                        if (!clanSrod5.Equals(""))
                        {
                            try
                            {
                                if (!listSrodstvo.Contains(Int32.Parse(clanSrod5)))
                                {
                                    string poruka = " Ne postoji srodstvo: " + clanSrod5;
                                    listaGresaka[brojGresaka] = poruka;
                                    brojGresaka++;
                                    // upisati u neki fajl greske
                                    fs1.WriteLine("Greska " + brojGresaka.ToString() + "-" + poruka + " u redu " + row);
                                    fs1.Flush();
                                }
                            }
                            catch (Exception ee)
                            {
                                string poruka = " Pogresno srodstvo: " + clanSrod5;
                                listaGresaka[brojGresaka] = poruka;
                                brojGresaka++;
                                // upisati u neki fajl greske
                                fs1.WriteLine("Greska " + brojGresaka.ToString() + "-" + poruka + " u redu " + row);
                                fs1.Flush();
                            }
                        }
                        }
                    catch (NullReferenceException ne)
                    {
                        string poruka = " Ne postoji srodstvo u dokumentu ";
                        listaGresaka[brojGresaka] = poruka;
                        brojGresaka++;
                        // upisati u neki fajl greske
                        fs1.WriteLine("Greska " + brojGresaka.ToString() + "-" + poruka + " u redu " + row);
                        fs1.Flush();
                    }
                    if (brojGresaka >= 20)
                    {
                        KillExcel(xlApp);
                        System.Threading.Thread.Sleep(500);
                        moveFile(fs1, par, nazivDat, true);
                        // Salji email
                        String imeGreske = ((FileStream)(fs1.BaseStream)).Name;
                        fs1.Close();
                        fs1.Dispose();
                        ispravna = false;
                  //      posaljiMail(fs, par, imeGreske, "Broj gresaka prelazi 20"); //"GRESKA TEST"); //listaGresaka.ToString());
                        return;
                        //   Environment.Exit(0);
                        //moveFile(fs, par, xd.naziv, true);
                    }

                    // provjera ukupnog iznosa
                    if (k == brojisplata)
                    {

                        if (!(Math.Round(double.Parse(ukupanIznosStavki.ToString()), 2) == (Math.Round(double.Parse(ukupanIznos.ToString()), 2)))) //Double.Parse(ukupanIznos)))
                        {
                            string poruka = " Ne slazu se isplate : " + ukupanIznos + "- " + ukupanIznosStavki;


                            listaGresaka[brojGresaka] = poruka;
                            brojGresaka++;
                            // upisati u neki fajl greske
                            fs1.WriteLine("Greska " + brojGresaka.ToString() + "-" + poruka);
                            fs1.Flush();
                        }
                        //broj stavki
                        if (!(Math.Round(double.Parse(ukupanBrojStavki.ToString()), 2) == (Math.Round(double.Parse(brojIsplata.ToString()), 2)))) //Double.Parse(ukupanIznos)))
                        {
                            string poruka = " Ne slaze se broj isplata : " + brojIsplata + "- " + ukupanBrojStavki;


                            listaGresaka[brojGresaka] = poruka;
                            brojGresaka++;
                            // upisati u neki fajl greske
                            fs1.WriteLine("Greska " + brojGresaka.ToString() + "-" + poruka);
                            fs1.Flush();
                        }
                        if (brojGresaka >= 1)
                        // na kraju se salje mail za manje od 20 gresaka
                        {
                            KillExcel(xlApp);
                            System.Threading.Thread.Sleep(500);
                            moveFile(fs, par, nazivDat, true);
                            String imeGreske = ((FileStream)(fs1.BaseStream)).Name;
                            fs1.Close();
                            fs1.Dispose();
                      //      posaljiMail(fs, par, imeGreske, "Manje od 20 gresaka"); //"GRESKA TEST"); //listaGresaka.ToString());
                            return;
                        }
                        else
                        {
                            KillExcel(xlApp);
                            System.Threading.Thread.Sleep(500);
                            moveFile(fs, par, nazivDat, false);
                            String imeGreske = ((FileStream)(fs1.BaseStream)).Name;
                            fs1.WriteLine("Broj isplata: " + brojIsplata.ToString() + " Iznos: " + ukupanIznos.ToString());
                            fs1.Flush();
                            fs1.WriteLine("Uspjesan prenos dokumenta " + nazivDat);
                            fs1.Close();
                            fs1.Dispose();
                      //      posaljiMail(fs, par, imeGreske, "Uspijesan import"); //"GRESKA TEST"); //listaGresaka.ToString());
                            return;
                        }
                    }
                }
                catch (Exception e)
                {
                    fs1.WriteLine(e.Message);
                    fs1.Flush();
                    //KillExcel(xlApp);
                    //moveFile(fs, par, nazivDat, true);
                }
            }
        }


        public void posaljiMail(StreamWriter fs, ParametriXml par, string porukaGreska, string subject)
        {
            MailMessage mail = new MailMessage();
            //            mail.From = new System.Net.Mail.MailAddress("bakir.zaciragic@gmail.com");
            mail.From = new System.Net.Mail.MailAddress(par.mailFrom);
            mail.Subject = subject;

            //create instance of smtpclient
            SmtpClient smtp = new SmtpClient();
            //            smtp.Port = 587;    //465;
            smtp.Port = par.smtpPort;    //465;
            smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtp.UseDefaultCredentials = false;
            //            smtp.Credentials = new NetworkCredential("bakir.zaciragic@gmail.com", "bakir2000");
            smtp.Credentials = new NetworkCredential(par.mailFrom, par.mailFromLozinka);
            //            smtp.Host = "smtp.gmail.com";
            smtp.Host = par.smtpHost;
            smtp.EnableSsl = true;
            //recipient address

            //            mail.To.Add(new MailAddress("zbakir@comp-it.ba"));
            //string mailTo;
            //string mailCc;
            foreach (string m in listeMailAdmin)
            {
                //                mail.To.Add(new MailAddress(eMail));
                mail.To.Add(new MailAddress(m));
            }
            foreach (string m in listeMail)
            {
                //                mail.To.Add(new MailAddress(eMail));
                mail.To.Add(new MailAddress(m));
            }
            //            mail.CC.Add(new MailAddress("zbakir@comp-2000.ba"));
            //          mail.CC.Add(new MailAddress(mailCc));
            //Formatted mail body
            mail.IsBodyHtml = true;
            string st = "U prilogu se nalaze informacije o importovanom fajlu koji je poslan sa lokacije " + porukaGreska;
            mail.Body = st;
            System.Net.Mail.Attachment attachment;
            attachment = new System.Net.Mail.Attachment(porukaGreska);
            mail.Attachments.Add(attachment);
            try
            {
                smtp.Send(mail);
            }
            catch (Exception e)
            {
                greskaTekst = "sentMail - " + e.Message;
                upisiGresku(listGresaka, greskaTekst);
                fs.WriteLine(greskaTekst);
                fs.Flush();
                greska = true;
            }
        }
        public void unesiIsplateUBazu(StreamWriter fs1, String nazivXmlDatoteke)
        {
            insertUBazu(fs1, objectArray, row, xlWorkbook, nazivXmlDatoteke);
            int ukupanBroj = Int32.Parse(brojIsplata); // objectArray.GetUpperBound(0) - 3;
            // procedura za obradu jednog dokumenta
            for (row = 4; row < 4 + ukupanBroj; row++)
               // for (int row = 4; row <= objectArray.GetUpperBound(0); row++)
            {
                try
                {
                    jmb = (objectArray[row, 2] == null) ? string.Empty : objectArray[row, 2].ToString();
                    prezime = (objectArray[row, 3] == null) ? string.Empty : objectArray[row, 3].ToString();
                    ime = (objectArray[row, 4] == null) ? string.Empty : objectArray[row, 4].ToString();
                    roditelj = (objectArray[row, 5] == null) ? string.Empty : objectArray[row, 5].ToString();
                    mjestoRodjenja = (objectArray[row, 6] == null) ? string.Empty : objectArray[row, 6].ToString();

          
                    drzavljanstvoId = (objectArray[row, 7] == null) ? string.Empty : objectArray[row, 7].ToString();
                    opcinaId = (objectArray[row, 8] == null) ? string.Empty : objectArray[row, 8].ToString();
                    idOpcine = dajOpcinu(opcinaId);
                    posta = (objectArray[row, 9] == null) ? string.Empty : objectArray[row, 9].ToString();
                    mjesto = (objectArray[row, 10] == null) ? string.Empty : objectArray[row, 10].ToString();
                    adresa = (objectArray[row, 11] == null) ? string.Empty : objectArray[row, 11].ToString();
          
                    pravoId = (objectArray[row, 12] == null) ? string.Empty : objectArray[row, 12].ToString();
                    iznos = (objectArray[row, 13] == null) ? string.Empty : objectArray[row, 13].ToString();
                    iznosPojedinacan = Double.Parse(iznos);


                    vrstaAktaId = (objectArray[row, 14] == null) ? string.Empty : objectArray[row, 14].ToString();
                    protokol = (objectArray[row, 15] == null) ? string.Empty : objectArray[row, 15].ToString();
                    DateTime datumAkta1 = Convert.ToDateTime((objectArray[row, 16] == null) ? string.Empty : objectArray[row, 16].ToString());
                    datumAkta = datumAkta1.ToString("yyyy-MM-dd").ToString();
                    odgovornoLice = (objectArray[row, 17] == null) ? string.Empty : objectArray[row, 17].ToString();
                    unosLice = (objectArray[row, 18] == null) ? string.Empty : objectArray[row, 18].ToString();



                    upisiListu(fs1);
                    // clanovi
                    //domacinstva
                    clanPrezime1 = (objectArray[row, 19] == null) ? string.Empty : objectArray[row, 19].ToString();
                    clanIme1 = (objectArray[row, 20] == null) ? string.Empty : objectArray[row, 20].ToString();
                    clanSrod1 = (objectArray[row, 21] == null) ? string.Empty : objectArray[row, 21].ToString();
                    clanGodiste1 = (objectArray[row, 22] == null) ? string.Empty : objectArray[row, 22].ToString();
                    vclanLK1 = (objectArray[row, 23] == null) ? string.Empty : objectArray[row, 23].ToString();

                    clanPrezime2 = (objectArray[row, 24] == null) ? string.Empty : objectArray[row, 24].ToString();
                    clanIme2 = (objectArray[row, 25] == null) ? string.Empty : objectArray[row, 25].ToString();
                    clanSrod2 = (objectArray[row, 26] == null) ? string.Empty : objectArray[row, 26].ToString();
                    clanGodiste2 = (objectArray[row, 27] == null) ? string.Empty : objectArray[row, 27].ToString();
                    vclanLK2 = (objectArray[row, 28] == null) ? string.Empty : objectArray[row, 28].ToString();

                    clanPrezime3 = (objectArray[row, 29] == null) ? string.Empty : objectArray[row, 29].ToString();
                    clanIme3 = (objectArray[row, 30] == null) ? string.Empty : objectArray[row, 30].ToString();
                    clanSrod3 = (objectArray[row, 31] == null) ? string.Empty : objectArray[row, 31].ToString();
                    clanGodiste3 = (objectArray[row, 32] == null) ? string.Empty : objectArray[row, 32].ToString();
                    vclanLK3 = (objectArray[row, 33] == null) ? string.Empty : objectArray[row, 33].ToString();

                    clanPrezime4 = (objectArray[row, 34] == null) ? string.Empty : objectArray[row, 34].ToString();
                    clanIme4 = (objectArray[row, 35] == null) ? string.Empty : objectArray[row, 35].ToString();
                    clanSrod4 = (objectArray[row, 36] == null) ? string.Empty : objectArray[row, 36].ToString();
                    clanGodiste4 = (objectArray[row, 37] == null) ? string.Empty : objectArray[row, 37].ToString();
                    vclanLK4 = (objectArray[row, 38] == null) ? string.Empty : objectArray[row, 38].ToString();

                    clanPrezime5 = (objectArray[row, 39] == null) ? string.Empty : objectArray[row, 39].ToString();
                    clanIme5 = (objectArray[row, 40] == null) ? string.Empty : objectArray[row, 40].ToString();
                    clanSrod5 = (objectArray[row, 41] == null) ? string.Empty : objectArray[row, 41].ToString();
                    clanGodiste5 = (objectArray[row, 42] == null) ? string.Empty : objectArray[row, 42].ToString();
                    vclanLK5 = (objectArray[row, 43] == null) ? string.Empty : objectArray[row, 43].ToString();

                    if (!clanSrod1.Equals(""))
                    {
                        upisiClan(fs1, clanPrezime1, clanIme1, clanSrod1, clanGodiste1, vclanLK1);
                    }

                    if (!clanSrod2.Equals(""))
                    {
                        upisiClan(fs1, clanPrezime2, clanIme2, clanSrod2, clanGodiste2, vclanLK2);
                    }
                    if (!clanSrod3.Equals(""))
                    {
                        upisiClan(fs1, clanPrezime3, clanIme3, clanSrod3, clanGodiste3, vclanLK3);
                    }
                    if (!clanSrod4.Equals(""))
                    {
                        upisiClan(fs1, clanPrezime4, clanIme4, clanSrod4, clanGodiste4, vclanLK4);
                    }
                    if (!clanSrod5.Equals(""))
                    {
                        upisiClan(fs1, clanPrezime5, clanIme5, clanSrod5, clanGodiste5, vclanLK5);
                    }

                }
                catch (Exception e)
                {
                    fs1.WriteLine("Greska - unesiIsplateUBazu "+ e.Message);
                }
            }
        }


        public void insertUBazu(StreamWriter fs1, object[,] excelNiz, int row, Workbook xlWorkbook1, String nazivXmlDatoteke)
        {
            unesiSpisak(fs1, (int)firmaId, Int32.Parse(excelNiz[2,3].ToString()), Int32.Parse(excelNiz[2, 4].ToString()), datumIsplate1, Int32.Parse(excelNiz[2, 5].ToString()), vrijeme);
            unesiUpload(fs1, nazivXmlDatoteke);

        }


        public void unesiSpisak(StreamWriter fs1, int Institucija, int PravniOsnov, int GrupaPrava, DateTime Datum, int ucestalostIsplata, DateTime vrijeme)
        {
            vrijeme = DateTime.Now;
            string komanda = "INSERT INTO dbo.Spiskovi(IsDeleted, InstitucijaId, PravniOsnovId, GrupaPravaId, " +
                              "Datum, ucestalostIsplataId, userId, LastModifiedTime) " +
                           "VALUES (@IsDeleted, @InstitucijaId, @PravniOsnovId, @GrupaPravaId, " +
                              "@Datum, @ucestalostIsplataId, @userId, @vrijeme)";

            try
            {
                connection.Open();
                SqlCommand command = new SqlCommand(komanda, connection);

                command.Parameters.AddWithValue("@IsDeleted", 0);
                command.Parameters.AddWithValue("@InstitucijaId", Institucija);
                command.Parameters.AddWithValue("@PravniOsnovId", PravniOsnov);
                command.Parameters.AddWithValue("@GrupaPravaId", GrupaPrava);
                //                command.Parameters.AddWithValue("@ucestalostIsplateId", institucija.ucestalostIsplate);
                command.Parameters.AddWithValue("@Datum", Datum);
                command.Parameters.AddWithValue("@ucestalostIsplataId", ucestalostIsplata);
                command.Parameters.AddWithValue("@vrijeme", vrijeme);
                command.Parameters.AddWithValue("@userId", userId);

                command.ExecuteNonQuery();

                SqlCommand getId = new SqlCommand("SELECT @@IDENTITY AS Id", connection);
                object id = getId.ExecuteScalar();
                if (id == null)
                {
                    spisakId = (int?)null;
                }
                else
                {
                    spisakId = Convert.ToInt32(id);
                }


                connection.Close();
            }
            catch (Exception e)
            {
                fs1.WriteLine("Greska kod unosa u bazu,  Spiskovi " + e.Message);
                fs1.Flush();
                connection.Close();
            }
        }

        public void unesiUpload(StreamWriter fs1, String nazivXmlDatoteke)
        {
            odgovor = "";
            vrijeme = DateTime.Now;
            string komanda = "INSERT INTO dbo.Upload(IsDeleted, userId, institucijaId, datoteka, vrijeme, odgovor, isOk, LastModifiedTime, spisakId, status) " +
                   "VALUES (@IsDeleted, @userId, @institucijaId, @datoteka, @vrijeme, @odgovor, @isOk, @vrijeme1, @spisakId, 'processing')";

            //            if ((string.IsNullOrEmpty(korisnik.jmb.Trim())))

            try
            {
                connection.Open();
                SqlCommand command = new SqlCommand(komanda, connection);

                command.Parameters.AddWithValue("@IsDeleted", 0);
                command.Parameters.AddWithValue("@userId", userId);
                command.Parameters.AddWithValue("@institucijaId", string.IsNullOrEmpty(firmaId.ToString()) ? Convert.DBNull : firmaId);
                //                string.IsNullOrEmpty(comment) ? (object)DBNull.Value : comment)
                command.Parameters.AddWithValue("@datoteka", nazivXmlDatoteke);
                command.Parameters.AddWithValue("@vrijeme", vrijeme);
                command.Parameters.AddWithValue("@odgovor", odgovor);
                command.Parameters.AddWithValue("@isOk", 1);
                command.Parameters.AddWithValue("@vrijeme1", vrijeme);
                command.Parameters.AddWithValue("@spisakId", spisakId);


                command.ExecuteNonQuery();

                SqlCommand getId = new SqlCommand("SELECT @@IDENTITY AS Id", connection);
                object id = getId.ExecuteScalar();

                if (id == null)
                {
                    uploadId = (int?)null;
                    //uploadIdStr = "null";
                }
                else
                {
                    uploadId = Convert.ToInt32(id);
                    // isOk = true;
                    // uploadIdStr = uploadId.ToString();
                }

                connection.Close();
            }
            catch (Exception e)
            {
                fs1.WriteLine("Greska kod unosa u bazu,  Upload " + e.Message);
                fs1.Flush();
                connection.Close();
               // isOk = false;
                //fs.Flush();
            }                 
        }

        void upisiListu(StreamWriter fs1)
        {
            bool isOk = true;
            vrijeme = DateTime.Now;
            odgovor = "";
            char spol;
            string muski = "01234";
            string zenski = "56789";
            string spolString = jmb.Substring(9, 1);
            if (muski.Contains(spolString))
            {
                spol = 'M';
            }
            else if (zenski.Contains(spolString))
            {
                spol = 'Z';
            }
            else
            {
                spol = 'N';
            }

            if (postojiLice(jmb) == false)
            {
                pisiNovoLice(fs1);
            }

          

            string komandaInsert =
                "INSERT INTO dbo.Liste(IsDeleted, SpisakId, Jmb, Prezime, Ime, Roditelj, MjestoRodjenja, " +
                    "Spol, Posta, Mjesto, Adresa, PravoId, OdgovornaOsoba, Obradjivac, Iznos, " +
                    "VrstaAktaId, DatumAkta, BrojAkta, DrzavaId, DPZId, liceId, LastModifiedTime, jmbStatusId) " +
                "VALUES (@IsDeleted, @SpisakId, @Jmb, @Prezime, @Ime, @Roditelj, @MjestoRodjenja, " +
                    "@Spol, @Posta, @Mjesto, @Adresa, @PravoId, @OdgovornaOsoba, @Obradjivac, @iznos, " +
                    "@VrstaAktaId, @DatumAkta, @BrojAkta, @DrzavaId, @DPZId, @liceId, @vrijeme, 5) ";                    
            try
            {
                connection.Open();
                SqlCommand command = new SqlCommand(komandaInsert, connection);

                command.Parameters.AddWithValue("@IsDeleted", 0);
                command.Parameters.AddWithValue("@spisakId", spisakId);
                command.Parameters.AddWithValue("@Jmb", jmb);
                command.Parameters.AddWithValue("@Prezime", prezime);
                command.Parameters.AddWithValue("@Ime", ime);
                if (roditelj != null)
                {
                    command.Parameters.AddWithValue("@Roditelj", roditelj);
                }
                else
                {
                    command.Parameters.AddWithValue("@Roditelj", DBNull.Value);
                }
                //                command.Parameters.AddWithValue("@Roditelj", korisnik.roditelj);
                if (mjestoRodjenja != null)
                {
                    command.Parameters.AddWithValue("@MjestoRodjenja", mjestoRodjenja);
                }
                else
                {
                    command.Parameters.AddWithValue("@MjestoRodjenja", DBNull.Value);
                }
                //                command.Parameters.AddWithValue("@MjestoRodjenja", korisnik.mjestoRodjenja);
                command.Parameters.AddWithValue("@Spol", spol);
                if (posta != null)
                {
                    command.Parameters.AddWithValue("@Posta", posta);
                }
                else
                {
                    command.Parameters.AddWithValue("@Posta", DBNull.Value);
                }
                //                command.Parameters.AddWithValue("@Posta", korisnik.posta);
                if (mjesto != null)
                {
                    command.Parameters.AddWithValue("@Mjesto", mjesto);
                }
                else
                {
                    command.Parameters.AddWithValue("@Mjesto", DBNull.Value);
                }
                //                command.Parameters.AddWithValue("@Mjesto", korisnik.mjesto);
                if (adresa != null)
                {
                    command.Parameters.AddWithValue("@Adresa", adresa);
                }
                else
                {
                    command.Parameters.AddWithValue("@Adresa", DBNull.Value);
                }
                //                command.Parameters.AddWithValue("@Adresa", korisnik.adresa);
                command.Parameters.AddWithValue("@PravoId", pravoId);
                if (odgovornoLice != null)
                {
                    command.Parameters.AddWithValue("@OdgovornaOsoba", odgovornoLice);
                }
                else
                {
                    command.Parameters.AddWithValue("@OdgovornaOsoba", DBNull.Value);
                }
                //                command.Parameters.AddWithValue("@OdgovornaOsoba", korisnik.odgovornoLice);
                if (unosLice != null)
                {
                    command.Parameters.AddWithValue("@Obradjivac", unosLice);
                }
                else
                {
                    command.Parameters.AddWithValue("@Obradjivac", DBNull.Value);
                }
                //                command.Parameters.AddWithValue("@Obradjivac", korisnik.unijeloLice);
                command.Parameters.AddWithValue("@Iznos", iznosPojedinacan);
                command.Parameters.AddWithValue("@VrstaAktaId", vrstaAktaId);
                if (datumAkta != null)
                {
                    command.Parameters.AddWithValue("@DatumAkta", datumAkta);
                }
                else
                {
                    command.Parameters.AddWithValue("@DatumAkta", DBNull.Value);
                }
                if (protokol != null)
                {
                    command.Parameters.AddWithValue("@BrojAkta", protokol);
                }
                else
                {
                    command.Parameters.AddWithValue("@BrojAkta", DBNull.Value);
                }
                //                command.Parameters.AddWithValue("@BrojAktaId", korisnik.akt);
                if (drzavljanstvoId != null)
                {
                    command.Parameters.AddWithValue("@DrzavaId", drzavljanstvoId);
                }
                else
                {
                    command.Parameters.AddWithValue("@DrzavaId", DBNull.Value);
                }
                //                command.Parameters.AddWithValue("@DrzavaId", korisnik.drzavljanstvo);
                if (idOpcine != null)
                {
                    command.Parameters.AddWithValue("@DPZId", idOpcine);
                }
                else
                {
                    command.Parameters.AddWithValue("@DPZId", DBNull.Value);
                }
                //                command.Parameters.AddWithValue("@DPZId", korisnik.opcinaStanovanja);
                command.Parameters.AddWithValue("@liceId", this.liceId);
                command.Parameters.AddWithValue("@vrijeme", vrijeme);

                command.ExecuteNonQuery();

                SqlCommand getId = new SqlCommand("SELECT @@IDENTITY AS Id", connection);
                object id = getId.ExecuteScalar();

                if (id == null)
                {
                    listaId = (int?)null;
                }
                else
                {
                    listaId = Convert.ToInt32(id);
                }

                connection.Close();
            }
            catch (Exception e)
            {

                fs1.WriteLine("Greska kod pisanja u bazu, lista " + e.Message);
                fs1.Flush();
                connection.Close();
                
            }
        }

        void pisiNovoLice(StreamWriter fs1)
        {
            vrijeme = DateTime.Now;
            bool greskaPisiLice = false;
          

            string komandaInsert = "INSERT INTO dbo.Lica (IsDeleted, Jmbg, Ime, Prezime, Roditelj, MjestoRodjenja," +
                        "DrzavaId, Posta, Mjesto, Adresa, DPZId, spol, LastModifiedTime) " +
                "VALUES (@IsDeleted, @Jmbg, @Ime, @Prezime, @Roditelj, @MjestoRodjenja, " +
                "@DrzavaId, @Posta, @Mjesto, @Adresa, @DPZId, @spol, @vrijeme)";

            char spol;
            string muski = "01234";
            string zenski = "56789";
            string oznakaPol = jmb.Substring(9, 1);
            if (muski.Contains(oznakaPol))
            {
                spol = 'M';
            }
            else if (zenski.Contains(oznakaPol))
            {
                spol = 'Z';
            }
            else
            {
                spol = 'N';
            }
            bool IsDeleted = false;
            //  Kontrola not null polja - (jmbg, ime, prezime, opcinaStanovanja, spol)
           
            try
            {
                connection.Open();
                SqlCommand command = new SqlCommand(komandaInsert, connection);

                command.Parameters.AddWithValue("@IsDeleted", 0);
                command.Parameters.AddWithValue("@Jmbg", jmb);
                command.Parameters.AddWithValue("@Prezime", prezime);
                command.Parameters.AddWithValue("@Ime", ime);
                command.Parameters.AddWithValue("@vrijeme", vrijeme);
                //               command.Parameters.AddWithValue("@spol", spol);
                if (roditelj != null)
                {
                    command.Parameters.AddWithValue("@Roditelj", roditelj);
                }
                else
                {
                    command.Parameters.AddWithValue("@Roditelj", DBNull.Value);
                }
                //                command.Parameters.AddWithValue("@Roditelj", korisnik.roditelj);
                if (mjestoRodjenja != null)
                {
                    command.Parameters.AddWithValue("@MjestoRodjenja", mjestoRodjenja);
                }
                else
                {
                    command.Parameters.AddWithValue("@MjestoRodjenja", DBNull.Value);
                }
                //                command.Parameters.AddWithValue("@MjestoRodjenja", korisnik.mjestoRodjenja);
                command.Parameters.AddWithValue("@Spol", spol);
                if (drzavljanstvoId != null)
                {
                    command.Parameters.AddWithValue("@DrzavaId", drzavljanstvoId);
                }
                else
                {
                    command.Parameters.AddWithValue("@DrzavaId", DBNull.Value);
                }
                //                command.Parameters.AddWithValue("@DrzavaId", korisnik.drzavljanstvo);
                if (posta != null)
                {
                    command.Parameters.AddWithValue("@Posta", posta);
                }
                else
                {
                    command.Parameters.AddWithValue("@Posta", DBNull.Value);
                }
                //                command.Parameters.AddWithValue("@Posta", korisnik.posta);
                if (mjesto != null)
                {
                    command.Parameters.AddWithValue("@Mjesto", mjesto);
                }
                else
                {
                    command.Parameters.AddWithValue("@Mjesto", DBNull.Value);
                }
                //                command.Parameters.AddWithValue("@Mjesto", korisnik.mjesto);
                if (adresa != null)
                {
                    command.Parameters.AddWithValue("@Adresa", adresa);
                }
                else
                {
                    command.Parameters.AddWithValue("@Adresa", DBNull.Value);
                }
                //                command.Parameters.AddWithValue("@Adresa", korisnik.adresa);
                if (idOpcine != null)
                {
                    command.Parameters.AddWithValue("@DPZId", idOpcine);
                }
                else
                {
                    command.Parameters.AddWithValue("@DPZId", DBNull.Value);
                }
                //                command.Parameters.AddWithValue("@DPZId", korisnik.opcinaStanovanja);

                command.ExecuteNonQuery();

                SqlCommand getId = new SqlCommand("SELECT @@IDENTITY AS Id", connection);
                object id = getId.ExecuteScalar();

                if (id == null)
                {
                    liceId = (int?)null;
                }
                else
                {
                    liceId = Convert.ToInt32(id);
                }

                connection.Close();
            }
            catch (Exception e)
            {
                fs1.WriteLine("Greska kod pisanja u bazu, lista " + e.Message);
                fs1.Flush();
                connection.Close();           
            }
        }


        void upisiClan(StreamWriter fs1, string prezime, string ime, string srodstvo, string godiste, string lk)
        {
         //   bool greskaPisiClan = false;
           
            vrijeme = DateTime.Now;

            string komanda = "INSERT INTO dbo.Domacinstva " +
                        "(IsDeleted, ListaId, Prezime, Ime, SrodstvoId, Godiste, BrojLK, LastModifiedTime) " +
                        "VALUES (@IsDeleted, @ListaId, @Prezime, @Ime, @SrodstvoId, @Godiste, @lk, @vrijeme)";

            //if ((string.IsNullOrEmpty(clan.ime.Trim())))
            //{
            //    greskaTekst = "Greska ime člana domaćinstva nije poznat - " + korisnik.jmb + ";" + korisnik.prezime + ";" + korisnik.ime;
            //    upisiGresku(listGresaka, greskaTekst);
            //    fs.WriteLine(greskaTekst);
            //    greska = true;
            //    greskaPisiClan = true;
            //}

            //if (greskaPisiClan == true)
            //{
            //    fs.Flush();
            //    return;
            //}

            try
            {
                connection.Open();
                SqlCommand command = new SqlCommand(komanda, connection);

                command.Parameters.AddWithValue("@IsDeleted", 0);
                command.Parameters.AddWithValue("@ListaId", listaId);
                command.Parameters.AddWithValue("@Prezime", prezime);
                command.Parameters.AddWithValue("@Ime", ime);
                command.Parameters.AddWithValue("@lk", lk);
                command.Parameters.AddWithValue("@vrijeme", vrijeme);
                if (srodstvo != null)
                {
                    command.Parameters.AddWithValue("@SrodstvoId", srodstvo);
                }
                else
                {
                    command.Parameters.AddWithValue("@SrodstvoId", DBNull.Value);
                }
                //                command.Parameters.AddWithValue("@SrodstvoId", clan.srodstvo);
                if (godiste != null)
                {
                    command.Parameters.AddWithValue("@Godiste", godiste);
                }
                else
                {
                    command.Parameters.AddWithValue("@Godiste", DBNull.Value);
                }
              
                command.ExecuteNonQuery();

                connection.Close();
            }
            catch (Exception e)
            {
                fs1.WriteLine("Greska kod pisanja u bazu, Domacinstva " + e.Message);
                fs1.Flush();
                connection.Close();
                
            }
        }
        public Boolean getUserId(StreamWriter fs, ParametriXml par, string userName)
        {
            bool korisnikPostoji = false;
            string komanda = "select id, eMail from dbo.users where username = @userName";
            string eMail = "";
            SqlCommand command;
            userId = (int?)null;
            string polje = "";
            //listeMailAdmin.Clear();
            try
            {
                connection.Open();
                command = new SqlCommand(komanda, connection);
                command.CommandType = System.Data.CommandType.Text;
                command.Parameters.AddWithValue("@userName", userName);

                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    polje = "id";
                    firmaId = reader.GetInt32(reader.GetOrdinal(polje));
                    userId = firmaId;
                    korisnikPostoji = true;
                    polje = "eMail";
                    if (!reader.IsDBNull(reader.GetOrdinal(polje)))
                    {
                        eMail = reader.GetString(reader.GetOrdinal(polje));
                        addEMail(listeMailAdmin, eMail, " - Korisnik");
                        //                        listeMailAdmin.Add(eMail);
                    }
                    else
                    {
                        eMail = string.Empty;
                    }
                }
            }
            catch (Exception e)
            {
                greskaTekst = "getUserId - Greska " + e.Message;
                upisiGresku(listGresaka, greskaTekst);
                fs.WriteLine(greskaTekst);
                greska = true;
                fs.Flush();
            }
            connection.Close();

            return korisnikPostoji;
        }


          public void moveFile(StreamWriter fs, ParametriXml par, string file, bool isErr)
        {
            string xmlJustFileName = Path.GetFileName(file);
            //buffer = Encoding.ASCII.GetBytes("Kopiranje datoteka: " + file + Environment.NewLine);
            //fs.Write(buffer, 0, buffer.Length);
            //buffer = Encoding.ASCII.GetBytes(Environment.NewLine);
            //fs.Write(buffer, 0, buffer.Length);
            //fs.Flush();
            string godinaErrStr;

            if (isErr)
            {
                godinaErrStr = godinaStr + "Err";
            }
            else
            {
                godinaErrStr = godinaStr;
            }

            string direktorij = Path.Combine(par.putArhiva, godinaErrStr);
            if (!Directory.Exists(direktorij))
            {
                Directory.CreateDirectory(direktorij);
            }
            direktorij = Path.Combine(direktorij, mjesecStr);
            if (!Directory.Exists(direktorij))
            {
                Directory.CreateDirectory(direktorij);
            }
            direktorij = Path.Combine(direktorij, danStr);
            if (!Directory.Exists(direktorij))
            {
                Directory.CreateDirectory(direktorij);
            }
            string porukaStr = "";
            try
            {
                string fileName = Path.GetFileName(file);
                string sourcePath = Path.GetDirectoryName(file);
                string targetPath = direktorij;
                string sourceFile = Path.Combine(sourcePath, fileName);
                string destFile = Path.Combine(targetPath, fileName);

                fs.WriteLine("     Move: " + " sourceFile: " + sourceFile + " TO destFile: " + destFile);
                fs.Flush();
                porukaStr = "File.Move(sourceFile, destFile)";
                File.Copy(sourceFile, destFile, true);
                File.Delete(sourceFile);
            }
            catch (Exception e)
            {
                fs.WriteLine("     Greska: Datoteka " + porukaStr + " >" + xmlJustFileName + "< " + e.Message);
                //greska = true;
               // greskaTekst = "     Greska: Datoteka " + porukaStr + " >" + xmlJustFileName + "< " + e.Message;
               // upisiGresku(listGresaka, greskaTekst);
               // fs.WriteLine(greskaTekst);
               fs.Flush();
                greska = true;
            }

        }
    }

}
