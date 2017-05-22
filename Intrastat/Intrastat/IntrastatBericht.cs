using System;
using System.Collections.Generic;
using System.Data.Odbc;
using AquaLight.Intrastat;
using System.Text.RegularExpressions;
using System.Xml;
using System.Text;
using System.IO;
using System.Xml.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using Aqualight;
using Aqualight.Intrastat;
using Aqualight.Database;
using System.Globalization;

namespace AquaLight.IntrastatBericht
{

    class IntrastatBericht
    {
        private String year;
        private String month;
        private String lastDay;
        private String databaseDSN;
        private String pathtoIntrastatBericht;        
        private List<BestellungElement> excelliste;
        private static object syncLock = new object();
        


        public IntrastatBericht(String year, String month, List<BestellungElement> list, String dsn, String pathtoIntrastatBericht, AuftragspositionenListe auftragspositionenliste)
        {
            this.year = year;
            this.month = month;
            this.excelliste = list;
            this.pathtoIntrastatBericht = pathtoIntrastatBericht;
            


            switch (month)
            {
                case "January": this.month = "01"; this.lastDay = "31"; break;
                case "February": this.month = "02"; this.lastDay = "28"; break;
                case "March": this.month = "03"; this.lastDay = "31"; break;
                case "April": this.month = "04"; this.lastDay = "30"; break;
                case "May": this.month = "05"; this.lastDay = "31"; break;
                case "June": this.month = "06"; this.lastDay = "30";  break;
                case "July": this.month = "07"; this.lastDay = "31"; break;
                case "August": this.month = "08"; this.lastDay = "31"; break;
                case "September": this.month = "09"; this.lastDay = "30"; break;
                case "October": this.month = "10"; this.lastDay = "31"; break;
                case "November": this.month = "11"; this.lastDay = "30"; break;
                case "December": this.month = "12"; this.lastDay = "31"; break;
            }
            this.databaseDSN = dsn;
            getIntrastatBericht();
        }

        public void getIntrastatBericht()
        {

            Auftragsliste auftragsliste = new Auftragsliste(year, month, lastDay);
            List<Auftrag> list = auftragsliste.List;                                   
            writeIntrastatXMLFile(list);            
        }

        private String transformSUCode(String sucode)
        {
            if (sucode.Equals("Stck"))
            {
                return "St";
            }
            if (sucode.Equals("kg"))
            {
                return "kg";
            }
            if (sucode.Equals("m"))
            {
                return "m";
            }
            return "St";
        }

        private void writeIntrastatXMLFile(List<Auftrag> auftraege)
        {
            DateTime dateTime = DateTime.Now;
            String date = dateTime.ToString("yyyy-MM-dd");
            String time = dateTime.ToString("HH:mm:ss");
            String fileName = "XGD98-" + year + month + "-" + dateTime.ToString("yyyyMMdd - HHmm");
            String period = year + "-" + month;
            String PSIId = "0367205016860001";
            String path = pathtoIntrastatBericht + "\\" + fileName + ".xml";


            using (XmlWriter writer = XmlWriter.Create(path))
            {
                lock (syncLock)
                {
                    writer.WriteStartDocument();
                    writer.WriteStartElement("INSTAT");
                    writer.WriteStartElement("Envelope");
                    writer.WriteElementString("envelopeId", "XGD98-" + year + month + "-" + DateTime.Now.Year + DateTime.Now.ToString("MM") + DateTime.Now.ToString("dd") + "-" + DateTime.Now.Hour + DateTime.Now.Minute);
                    writer.WriteStartElement("DateTime");
                    writer.WriteElementString("date", date);
                    writer.WriteElementString("time", time);
                    writer.WriteEndElement();
                    writer.WriteStartElement("Party");
                    writer.WriteAttributeString("partyRole", "receiver");
                    writer.WriteAttributeString("partyType", "CC");                    
                    writer.WriteElementString("partyId", "00");
                    writer.WriteElementString("partyName", "Statistisches Bundesamt");
                    writer.WriteStartElement("Address");
                    writer.WriteElementString("streetName", "Gustav-Stresemann-Ring 11");
                    writer.WriteElementString("postalCode", "65189");
                    writer.WriteElementString("cityName", "Wiesbaden");
                    writer.WriteEndElement(); //Address End
                    writer.WriteEndElement();//Party End
                    writer.WriteStartElement("Party");
                    writer.WriteAttributeString("partyRole", "sender");
                    writer.WriteAttributeString("partyType", "PSI"); //korrekt
 
                    writer.WriteElementString("partyId", PSIId);
                    writer.WriteElementString("partyName", "Aqua Light GmbH");
                    writer.WriteElementString("interchangeAgreementId", "XGD98");
                    writer.WriteStartElement("Address");
                    writer.WriteElementString("streetName", "Am Basterpohl");
                    writer.WriteElementString("streetNumber", "16");
                    writer.WriteElementString("postalCode", "49565");
                    writer.WriteElementString("cityName", "Bramsche");
                    writer.WriteElementString("countryName", "Deutschland");
                    writer.WriteElementString("phoneNumber", "05468 939446");
                    writer.WriteElementString("faxNumber", "05469 939153");
                    writer.WriteElementString("e-mail", "info@aqualight.de");
                    writer.WriteElementString("URL", "https://aqualight.de");                   
                    writer.WriteEndElement(); //ADDress End
                    writer.WriteEndElement(); //Party End
                    writer.WriteElementString("testIndicator", "false");                     //Test
                    writer.WriteElementString("softwareUsed", "Aqualight Warehouse System");

                    int i = 1;
                    int previousDeclaration = 0;

                    foreach (Auftrag auftrag in auftraege)
                    {
                        previousDeclaration = i - 1;


                        writer.WriteStartElement("Declaration");
                        writer.WriteElementString("declarationId", i.ToString());
                        writer.WriteStartElement("DateTime");
                        writer.WriteElementString("date", date);
                        writer.WriteElementString("time", time);
                        writer.WriteEndElement();
                        writer.WriteElementString("referencePeriod", period);
                        writer.WriteElementString("PSIId", PSIId);                              //zwingend: Identifikator des Melders
                        writer.WriteStartElement("Function");
                        writer.WriteElementString("functionCode", "O");
                        if (previousDeclaration < 1)
                        {
                            writer.WriteElementString("previousDeclarationId", "");
                        }
                        else
                        {
                            writer.WriteElementString("previousDeclarationId", previousDeclaration.ToString());
                        }
                        writer.WriteEndElement();
                        writer.WriteElementString("declarationTypeCode", "");
                        writer.WriteElementString("flowCode", "D");                             //Flowcode D = Versand, A = Eingang
                        writer.WriteElementString("currencyCode", "2");                         //Währungscode 2 = Euro

                        
                        int j = 1;
                        foreach (Auftragsposition pos in auftrag.auftragspositionenListe)
                        {
                            


                            if (!String.IsNullOrEmpty(pos.Anzahl) & Int32.Parse(pos.Anzahl) > 0)
                            {
                                if (String.IsNullOrEmpty(pos.Gewicht))
                                {
                                    pos.Gewicht = "1"; //Falls kein Gewicht da ist, nehme 1kg an.
                                }

                                if (String.IsNullOrEmpty(pos.Zoll))
                                {
                                    pos.Zoll = "84219900"; //Falls keine Zollnummer vorhanden ist, dann wähle diese Zollnummer
                                }


                                //Manchmal fehlt einfach die Null vorne...
                                if (Regex.IsMatch(pos.Gesamtpreis, "^'.'"))
                                {
                                    pos.Gesamtpreis = "0" + pos.Gesamtpreis;
                                }
                                
                                double gesamtpreis = double.Parse(pos.Gesamtpreis, CultureInfo.InvariantCulture);
                                //Preis aufrunden
                                if (gesamtpreis > 0.0)
                                {
                                    gesamtpreis = Math.Ceiling(gesamtpreis);
                                    pos.Gesamtpreis = gesamtpreis.ToString();
                                }


                                string[] gewicht2 = pos.Gewicht.Split(',');
                                if (gewicht2.Length > 0)
                                {
                                    pos.Gewicht = gewicht2[0];
                                }

                                String model = pos.Artikelnummer;
                                String bezeichnung = pos.Bezeichnung;
                                //Fehlerbereinigung der Falscheingabe durch Mitarbeiter
                                if (model.Equals("DIV-00098", StringComparison.OrdinalIgnoreCase))
                                {
                                    pos.Anzahl = (int.Parse(pos.Anzahl) / 1000).ToString();
                                    pos.Gewicht = (int.Parse(pos.Gewicht) / 1000).ToString();
                                }
                                if (bezeichnung.Equals("Fischtransportbeutel 225x500mm / 80µm", StringComparison.OrdinalIgnoreCase))
                                {
                                    pos.Gewicht = (int.Parse(pos.Gewicht) / 1000).ToString();
                                }
                                if (bezeichnung.Equals("Schlauch, Silikon 6/8mm - pro m", StringComparison.OrdinalIgnoreCase))
                                {
                                    pos.Gewicht = (int.Parse(pos.Gewicht) / 1000).ToString();
                                    pos.Anzahl = (int.Parse(pos.Anzahl) / 10).ToString();
                                }
                                //Falls etwas mit der Anzahl der Artikel nicht stimmt
                                if (int.Parse(pos.Anzahl) == 0)
                                {

                                    if (bezeichnung.Equals("Aktivkohle, pro kg (25kg/Sack)", StringComparison.OrdinalIgnoreCase))
                                    {
                                        pos.Anzahl = (int.Parse(pos.Gewicht) / 1.04).ToString();
                                    }
                                    if (int.Parse(pos.Anzahl) == 0)
                                    {
                                        pos.Anzahl = "1";
                                    }
                                }
                                if(int.Parse(pos.Gewicht) > 100000)
                                {
                                    pos.Gewicht = (int.Parse(pos.Gewicht) / 1000).ToString();
                                }
                                if(bezeichnung.Equals("Calciumhydroxid, per kg - VE:25kg"))
                                {
                                    pos.Anzahl = ((int.Parse(pos.Anzahl)) / 100).ToString();
                                     pos.Gewicht = int.Parse(pos.Anzahl).ToString();
                                }

                                if (model.Equals("AQUALIGHT20000", StringComparison.OrdinalIgnoreCase))
                                {
                                    pos.Gewicht = pos.Anzahl;
                                }
                                if (int.Parse(pos.Gewicht) == 0)
                                {
                                    pos.Gewicht = "1";
                                }



                                writer.WriteStartElement("Item");
                                writer.WriteElementString("itemNumber", j.ToString());
                                writer.WriteStartElement("CN8");
                                writer.WriteElementString("CN8Code", pos.Zoll);
                                writer.WriteElementString("SUCode", transformSUCode(pos.Mengeneinheit));  //IM SVZ doppelt prüfen
                                writer.WriteEndElement(); //CN8
                                writer.WriteElementString("goodsDescription", pos.Bezeichnung);
                                writer.WriteElementString("MSConsDestCode", auftrag.Laendercode);
                                writer.WriteElementString("quantityInSU", pos.Anzahl);
                                writer.WriteElementString("netMass", pos.Gewicht);
                                writer.WriteElementString("invoicedAmount", pos.Gesamtpreis);
                                writer.WriteElementString("invoiceNumber", auftrag.auftragsnummer);
                                writer.WriteElementString("statisticalValue", pos.Gesamtpreis);   //Im SVZ doppelt prüfen
                                writer.WriteStartElement("NatureOfTransaction");
                                writer.WriteElementString("natureOfTransactionACode", "1");
                                writer.WriteElementString("natureOfTransactionBCode", "1");
                                writer.WriteEndElement(); //Nature of Transaction
                                writer.WriteElementString("modeOfTransportCode", "3"); //Im SVZ doppelt prüfen 3 bedeutet per Straßenverkehr
                                writer.WriteElementString("regionCode", "03");         //Im SVZ doppelt prüfen 03 bedeutet Niedersachsern
                                writer.WriteEndElement(); //Item
                                j++;
                            }
                        }
                            
                    
                    writer.WriteEndElement(); //Declaration
                    i++;
                }

                   
                    writer.WriteEndElement();



                    writer.WriteEndElement();


                    writer.WriteEndDocument();
                    writer.Close();

                    
                    Encoding ansi = Encoding.GetEncoding(1250);
                    
                    string xml = File.ReadAllText(path);

                    XDocument xmlDoc = XDocument.Parse(xml);

                    File.WriteAllText(
                        path,
                        @"<?xml version=""1.0"" encoding=""ISO-8859-1""?>" + xmlDoc.ToString(),
                       ansi
                    );
                }
            }
        }

    }
}
