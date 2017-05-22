using AquaLight.Intrastat;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading.Tasks;

namespace Aqualight.Intrastat
{
    class Auftragsliste
    {

        private List<Auftrag> list = null;
        private String table = "AuftraegeArchiv__2004";

        /// <summary>
        /// Getter und Setter für die Auftragsliste
        /// </summary>
        internal List<Auftrag> List
        {
            get
            {
                return list;
            }

            set
            {
                list = value;
            }
        }
        /// <summary>
        /// Erstellt selbständig die Auftragsliste und speichert diese in eine Datei zum cachen
        /// </summary>
        /// <param name="year">Jahr der Aufträge</param>
        /// <param name="month">Monat der Aufträge</param>
        /// <param name="lastDay">Letzter Tag im Monat</param>
        public Auftragsliste(String year, String month, String lastDay) {
            String AuftragBinaryPath = "Auftragsliste" + month + year + ".bin";

            try
            {
                IFormatter formatter = new BinaryFormatter();
                Stream stream = new FileStream(AuftragBinaryPath, FileMode.Open, FileAccess.Read, FileShare.Read);
                if (stream.Length > 0)
                {
                    List = (List<Auftrag>)formatter.Deserialize(stream);
                }
                stream.Close();
            }
            catch(Exception e)
            {               
                List = getAuftragListe(year, month, lastDay);           
                IFormatter formatter = new BinaryFormatter();
                Stream stream = new FileStream(AuftragBinaryPath, FileMode.Create, FileAccess.Write, FileShare.None);
                formatter.Serialize(stream, List);
                stream.Close();
            }
            if(List != null)
            {
                List = addAuftragspositionenToList(List);
            }
        }
        /// <summary>
        /// Holt die Auftragsliste aus der ODBC Datenquelle
        /// </summary>
        /// <param name="year">Jahr der Aufträge</param>
        /// <param name="month">Monat der Aufträge</param>
        /// <param name="lastDay">Letzter Tag des Monats (Wichtig um Datenbankabfragen korrekt zu halten)</param>
        /// <returns></returns>
        private List<Auftrag> getAuftragListe(String year, String month, String lastDay)
        {
            List<Auftrag> list = new List<Auftrag>();
            
            OdbcConnection database = Database.Database.getDatabase();
            OdbcDataReader DbReader = null;

            int fCount = 0;
            try
            {
                database.Open();
                OdbcCommand DbCommand = database.CreateCommand();
                DbCommand.CommandText = "SELECT DISTINCT Auftrag, Laendercode, Datum, Bruttosumme, Gesamtgewicht FROM "+table+" WHERE Datum >= {d '" + year + "-" + month + "-01' } AND Datum <= {d '" + year + "-" + month + "-" + lastDay + "' } AND (Laendercode = 'AT' OR Laendercode = 'BE' OR Laendercode = 'BG' OR Laendercode = 'CY' OR Laendercode = 'CZ' OR Laendercode = 'DK' OR Laendercode = 'EE' OR Laendercode = 'ES' OR Laendercode = 'FI' OR Laendercode = 'FR' OR Laendercode = 'GB' OR Laendercode = 'GR' OR Laendercode = 'HU' OR Laendercode = 'IE' OR Laendercode = 'IT' OR Laendercode = 'LT' OR Laendercode = 'LU' OR Laendercode = 'LV' OR Laendercode = 'MT' OR Laendercode = 'NL' OR Laendercode = 'PL' OR Laendercode = 'PT' OR Laendercode = 'RO' OR Laendercode = 'SE' OR Laendercode = 'SI' OR Laendercode = 'SK') AND Bruttosumme > 0";
                DbReader = DbCommand.ExecuteReader();
                fCount = DbReader.FieldCount;
            }
            catch (Exception e)
            {
                Console.WriteLine("Fehler bei der Datenbankverbindung. " + e.ToString());
            }
            if (DbReader != null)
            {
                while (DbReader.Read())
                {
                    Auftrag auftrag = new Auftrag();

                    for (int i = 0; i < fCount; i++)
                    {
                        String col = DbReader.GetString(i);
                        switch (i)
                        {
                            case 0:
                                auftrag.auftragsnummer = col;
                                break;
                            case 1: auftrag.Laendercode = col; break;
                            case 2: auftrag.date = col; break;
                            case 3: auftrag.bruttowert = col; break;
                            case 4: auftrag.gewicht = col; break;
                        }
                    }

                    Boolean doppelt = false;
                    //check ob Auftrag doppelt                        
                    foreach (Auftrag a in list)
                    {

                        //falls Auftrag doppelt ist tue nichts
                        if (a.auftragsnummer.Equals(auftrag.auftragsnummer))
                        {
                            doppelt = true;        
                        }
                    }
                    if (!doppelt) {
                        list.Add(auftrag);
                    }

                }
            }
            if(list.Count == 0)
            {
                if (table.Equals("AuftraegeArchiv__2004"))
                {
                    table = "Auftraege";
                    return getAuftragListe(year, month, lastDay);
                }
                else return list;
            }
            return list;
        }
        /// <summary>
        /// Fügt die Auftragspositionen Liste zu dem Auftrag hinzu.
        /// </summary>
        /// <param name="list">Liste aller Aufträge</param>
        /// <returns>Liste aller Aufträge mit Auftragspositionen</returns>
        private List<Auftrag> addAuftragspositionenToList(List<Auftrag> list)
        {
            AuftragspositionenListe elem = new AuftragspositionenListe();
            List<Auftragsposition> positionenListe = elem.List;

            foreach(Auftrag auftrag in list)
            {
                List<Auftragsposition> erg = new List<Auftragsposition>();
                foreach(Auftragsposition pos in positionenListe)
                {
                    if(String.Compare(pos.Vorgangsnummer, auftrag.auftragsnummer) == 0)
                    {
                        if (checkPosition(pos))
                        {
                            
                            erg.Add(addExcelDataToPosition(pos));
                        }
                       
                    }
                }
                if (checkAuftrag(auftrag))
                {
                    //Diese Implementierung, da das Vergleichen der vollen Listen zu lange dauert
                    //Elemente, die in der Stückliste vorkommen löschen und hinzufügen
                    foreach (Auftragsposition af in erg)
                    {
                        if (af.Stuecklistenartikel)
                        {
                            foreach(Auftragsposition artikel in erg)
                            {
                                if (!artikel.Stuecklistenartikel)
                                {
                                    if (artikel.Artikelnummer.Equals(af.Artikelnummer))
                                    {
                                        //Doppelte Einträge von Stücklisten items rauslöschen
                                        erg.Remove(artikel);
                                    }
                                }
                            }
                        }
                        
                    }
                    auftrag.auftragspositionenListe = erg;
                }
                
            }
           
            return list;
        }
        /// <summary>
        /// Artikeldaten durch Excelliste ergänzen
        /// </summary>
        /// <param name="pos"></param>
        /// <returns></returns>
        private Auftragsposition addExcelDataToPosition(Auftragsposition pos)
        {
            List<BestellungElement> list = ExcelParser.ExcelParser.getExcelParser();
            foreach (BestellungElement e in list)
            {
                if (String.Compare(e.Artikelnummer, pos.Artikelnummer, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    pos.Zoll = e.Zollnummer;
                    if (pos.Anzahl != null)
                    {
                        if (String.IsNullOrEmpty(pos.Gewicht))
                        {
                            pos.Gewicht = (int.Parse(pos.Anzahl) * e.Gewicht).ToString();
}
                        else
                        {
                            pos.Gewicht = (int.Parse(pos.Anzahl) * double.Parse(pos.Gewicht)).ToString();
                        }
                    }
                    else
                    {
                        pos.Gewicht = e.Gewicht.ToString();
                        pos.Anzahl = "1";
                    }
                    
                    break; //In foreach-Schleife gefunden
                }
            }


            return pos;
        }

        /// <summary>
        /// Prüft ob die Auftragsposition eine sinnvolle Position für Intrastat ist
        /// </summary>
        /// <param name="pos">Auftragsposition</param>
        /// <returns>true falls ok, false falls nicht</returns>
        private bool checkPosition(Auftragsposition pos)
        {
            String model = pos.Artikelnummer;

            bool ok = true;
            
            model = model.Trim();
            if (model.Equals("tv", StringComparison.OrdinalIgnoreCase))
            {
                ok = false;
            }
            if (model.Equals("k", StringComparison.OrdinalIgnoreCase))
            {
                ok = false;
            }
            if (model.Equals("tv", StringComparison.OrdinalIgnoreCase))
            {
                ok = false;
            }
            if (model.Equals("t", StringComparison.OrdinalIgnoreCase))
            {
                ok = false;
            }
            if (model.Equals("z", StringComparison.OrdinalIgnoreCase))
            {
                ok = false;
            }
            if (model.Equals("tl", StringComparison.OrdinalIgnoreCase))
            {
                ok = false;
            }
            if (model.Equals("0", StringComparison.OrdinalIgnoreCase))
            {
                ok = false;
            }
            if (model.Equals("kv", StringComparison.OrdinalIgnoreCase))
            {
                ok = false;
            }
            if (model.Equals("1", StringComparison.OrdinalIgnoreCase))
            {
                ok = false;
            }
            if (model.Equals("e", StringComparison.OrdinalIgnoreCase))
            {
                ok = false;
            }
            if (model.Equals("tve", StringComparison.OrdinalIgnoreCase))
            {
                ok = false;
            }
            if (model.Equals("cd", StringComparison.OrdinalIgnoreCase))
            {
                ok = false;
            }
            if (model.Equals("m", StringComparison.OrdinalIgnoreCase))
            {
                ok = false;
            }
            if (model.Equals("n", StringComparison.OrdinalIgnoreCase))
            {
                ok = false;
            }
            if (model.Equals("tp", StringComparison.OrdinalIgnoreCase))
            {
                ok = false;
            }
            if (model.Equals("Z1", StringComparison.OrdinalIgnoreCase))
            {
                ok = false;
            }
            if (model.Equals("Z2", StringComparison.OrdinalIgnoreCase))
            {
                ok = false;
            }
            if (model.Equals("vf", StringComparison.OrdinalIgnoreCase))
            {
                ok = false;
            }
            if (model.Equals("v", StringComparison.OrdinalIgnoreCase))
            {
                ok = false;
            }
            return ok;
        }
        /// <summary>
        /// Prüft ob der Auftrag ok ist.
        /// </summary>
        /// <param name="auftrag">Auftrag</param>
        /// <returns>true falls der Auftrag ok ist, false falls nicht</returns>
        private bool checkAuftrag(Auftrag auftrag)
        {
            bool success = true;
            if (auftrag.auftragspositionenListe != null)
            {
                if (auftrag.auftragspositionenListe.Count == 0)
                {
                    success = false;
                }

                if (auftrag.auftragspositionenListe.Count == 1)
                {
                    Auftragsposition pos = auftrag.auftragspositionenListe[0];
                    if (String.Compare(pos.Artikelnummer, "tv", true) == 0)
                    {
                        success = false;
                    }
                }
            }
            return success;
        }
    }
}
