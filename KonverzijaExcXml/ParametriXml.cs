using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace KonverzijaExcXml
{
        public class ParametriXml
        {
            public string urlDb;
            public string bazaDb;
            public string korisnikDb;
            public string lozinkaDb;
            public string xsdFile;
            public string putConfig;
            public string fileConfig;
            public string putLog;
            public string putanja;
            public string putTmpXml;
            public string format;
            public string putArhiva;
            public int brojDatoteka = 999999;
            public bool isTrace = false;
            public string traceStr;
            public string mailFrom;
            public string mailAdmin = "";
            public string mailFromLozinka;
            public int smtpPort;
            public string smtpHost;
            public bool isTest = false;
            public string testStr;
            public string testTekst = "";
            public long fileSize = 50000;

            XmlDocument doc = new XmlDocument();
            string nodeName;
            string getSadrzaj(string tekst)
            {
                int zadnji;
                int prvi;
                string ch = tekst.Substring(tekst.Length - 1, 1);

                if (!int.TryParse(ch, out zadnji))
                {
                    zadnji = -1;
                }

                if (zadnji >= 0)
                {
                    tekst = tekst.Substring(0, tekst.Length - 1);
                }

                ch = tekst.Substring(tekst.Length - 1, 1);

                if (!int.TryParse(ch, out prvi))
                {
                    prvi = -1;
                }

                if (prvi >= 0)
                {
                    tekst = tekst.Substring(0, tekst.Length - 1);
                }

                tekst = tekst.Substring(prvi, tekst.Length - prvi - zadnji);
                return tekst;
            }
            //        public void citajXml(FileStream fs)
            public void citajXml()
            {
                //                string datoteke = "c://temp//imp//config.xml";
                var file = Path.Combine(putConfig, "config.xml");
                fileConfig = file;
                string tekst = File.ReadAllText(file);
                doc.LoadXml(tekst);
                XmlNode root = doc.FirstChild;

                for (int i = 0; i < root.ChildNodes.Count; i++)
                {
                    nodeName = root.ChildNodes[i].Name;
                    if (nodeName.ToLower() == "lozinkadb")
                    {
                        this.lozinkaDb = root.ChildNodes[i].InnerText;
                        this.lozinkaDb = getSadrzaj(this.lozinkaDb);
                    }
                    else if (nodeName.ToLower() == "korisnikdb")
                    {
                        this.korisnikDb = root.ChildNodes[i].InnerText;
                        this.korisnikDb = getSadrzaj(this.korisnikDb);
                    }
                    else if (nodeName.ToLower() == "urldb")
                    {
                        this.urlDb = root.ChildNodes[i].InnerText;
                    }
                    else if (nodeName.ToLower() == "xsd")
                    {
                        this.xsdFile = root.ChildNodes[i].InnerText;
                    }
                    else if (nodeName.ToLower() == "bazadb")
                    {
                        this.bazaDb = root.ChildNodes[i].InnerText;
                    }
                    else if (nodeName.ToLower() == "log")
                    {
                        this.putLog = root.ChildNodes[i].InnerText;
                    }
                    else if (nodeName.ToLower() == "arhiva")
                    {
                        this.putArhiva = root.ChildNodes[i].InnerText;
                    }
                    else if (nodeName.ToLower() == "xmltmp")
                    {
                        this.putTmpXml = root.ChildNodes[i].InnerText;
                    }
                    else if (nodeName.ToLower() == "brojdatoteka")
                    {
                        this.brojDatoteka = Convert.ToInt32(root.ChildNodes[i].InnerText);
                    }
                    else if (nodeName.ToLower() == "filesize")
                    {
                        this.fileSize = Convert.ToInt64(root.ChildNodes[i].InnerText);
                    }
                    else if (nodeName.ToLower() == "mailfrom")
                    {
                        this.mailFrom = root.ChildNodes[i].InnerText;
                    }
                    else if (nodeName.ToLower() == "mailadmin")
                    {
                        this.mailAdmin = root.ChildNodes[i].InnerText;
                    }
                    else if (nodeName.ToLower() == "mailfromlozinka")
                    {
                        this.mailFromLozinka = root.ChildNodes[i].InnerText;
                        this.mailFromLozinka = getSadrzaj(this.mailFromLozinka);
                    }
                else if (nodeName.ToLower() == "putanja")
                {
                    this.putanja = root.ChildNodes[i].InnerText;
                }
                else if (nodeName.ToLower() == "smtpport")
                    {
                        try
                        {
                            this.smtpPort = Convert.ToInt32(root.ChildNodes[i].InnerText);
                            //                    this.mailFromLozinka = getSadrzaj(this.mailFromLozinka);
                        }
                        catch
                        {

                        }
                    }
                    else if (nodeName.ToLower() == "smtphost")
                    {
                        this.smtpHost = root.ChildNodes[i].InnerText;
                        //                    this.mailFromLozinka = getSadrzaj(this.mailFromLozinka);
                    }
                    else if (nodeName.ToLower() == "trace")
                    {
                        this.traceStr = root.ChildNodes[i].InnerText.ToLower();
                        if (this.traceStr.Equals("true"))
                        {
                            this.isTrace = true;
                        }
                        else
                        {
                            this.isTrace = false;
                        }
                    }
                    else if (nodeName.ToLower() == "test")
                    {
                        this.testStr = root.ChildNodes[i].InnerText.ToLower();
                        if (this.testStr.ToLower().Equals("da"))
                        {
                            this.isTest = true;
                            testTekst = "Test način rada! ";
                        }
                        else
                        {
                            this.isTest = false;
                            testTekst = "Import - ";
                        }
                    }
                    //else
                    //{
                    //    greska = greska + "Pogresa tag" + root.ChildNodes[i].InnerText + "; ";
                    //    //fConfig.Write(buffer, 0, buffer.Length);
                    //}
                }
                //                    pisiDBParametreToFs(fs);
            }
           
            public void pisiDBParametreToFs(StreamWriter fs)
            {
                fs.WriteLine("Config: " + fileConfig);
                fs.WriteLine();
                fs.WriteLine("IMP - parametri");
                fs.WriteLine();
                fs.WriteLine("DB ");
                fs.WriteLine("UrlDb: " + this.urlDb);
                fs.WriteLine("Baza: " + this.bazaDb);
                fs.WriteLine("KorisnikDb: " + this.korisnikDb);
                fs.WriteLine("LozinkaDb: " + this.lozinkaDb);
                fs.WriteLine();
                fs.WriteLine("XSD: " + this.xsdFile);
                fs.WriteLine("Log: " + this.putLog);
                fs.WriteLine("XML: " + this.putTmpXml);
                fs.WriteLine("Broj datoteka: " + this.brojDatoteka.ToString());
                fs.WriteLine("Arhiva: " + this.putArhiva);
                fs.WriteLine();
                fs.WriteLine("MailFrom: " + this.mailFrom);
                fs.WriteLine("MailFromLozinka: " + this.mailFromLozinka);
                fs.WriteLine("MailSmtp: " + this.smtpPort);
                fs.WriteLine("MailSmtpHost: " + this.smtpHost);
                fs.WriteLine();
                fs.WriteLine("MailAdmin: " + this.mailAdmin);
                fs.WriteLine();
                fs.WriteLine("isTest:" + this.isTest.ToString());
                fs.WriteLine("isTrace: " + this.isTrace.ToString());
                fs.WriteLine();
                fs.Flush();
            }
        }
 }



