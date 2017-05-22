using AquaLight.Intrastat;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Aqualight.Intrastat
{
    [Serializable]
    class AuftragspositionenListe
    {
        private List<Auftragsposition> list = new List<Auftragsposition>();
        private String fileName = "AuftragspositionenListe.bin"; 

        public AuftragspositionenListe()
        {
            //Kann sein, dass die Liste noch im Hauptspeicher liegt. - Kann nicht sein, wird vom Garbage Collector vorher abgeräumt
             
                try
                {
                    IFormatter formatter = new BinaryFormatter();
                    Stream stream = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

                    if (stream.Length > 0 & !isOutdated(fileName))
                    {
                        List = (List<Auftragsposition>)formatter.Deserialize(stream);
                    }
                    stream.Close();
                }
                catch (FileNotFoundException fnfe)
                {
                    System.Console.WriteLine("Auftragspositionenliste wurde nicht gefunden. " + fnfe.ToString());
                }
                if (List.Count == 0)
                {
                    getAuftragspositionenListe("AuftraegeArchivPos__2004", "AuftraegeArchivPos_Stuecklisten__2004");
                    getAuftragspositionenListe("AuftraegePositionen_Artikel", "AuftraegePositionen_Stuecklisten");
                    //Speichern aller Daten in einem Element
                    IFormatter formatter = new BinaryFormatter();
                    Stream stream = new FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.ReadWrite);
                    formatter.Serialize(stream, List);
                    stream.Flush();
                    stream.Close();
                }
            
        }

        internal List<Auftragsposition> List
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
        /// Schaut, ob die Positionenliste schon 1 Monat alt ist. Falls ja, dann wird diese erneuert.
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private bool isOutdated(String fileName)
        {
            try
            {
                DateTime access = File.GetLastWriteTime(fileName);
                if (access < DateTime.Now.AddMonths(-1) | access.Month != DateTime.Now.Month)
                    return true;
            }
            catch (Exception e)
            {
                Console.WriteLine("Info: Auftragspositionenbinary nicht gefunden -  " + e.Message);
            }
            return false;
        }


        private void getAuftragspositionenListe(String ArtikelTabelle, String StueckListenTabelle)
        {

            OdbcConnection database = Database.Database.getDatabase();
            
            OdbcCommand DbCommand = database.CreateCommand();
            OdbcDataReader DbReader = null;
            int fCount = 0;

            DbCommand.CommandText = "SELECT Vorgangsnummer, Artikelnummer, Gesamtpreis, Bezeichnung1, Menge, Mengeneinheit FROM "+ArtikelTabelle;
            try
            {
                DbReader = DbCommand.ExecuteReader();
                fCount = DbReader.FieldCount;
            }
            catch (Exception e)
            {
                //Falls die Datenbank nicht aufrufbar ist
                try
                {
                    database.Open();
                    DbReader = DbCommand.ExecuteReader();
                    fCount = DbReader.FieldCount;
                }
                catch (Exception exc)
                {
                    Console.WriteLine("Could not restore database." + exc.ToString());
                }
            }

            //Artikeldaten durchgehen            
            while (DbReader.Read())
            {
                Auftragsposition auftragsposition = new Auftragsposition();
                for (int i = 0; i < fCount; i++)
                {
                    String col = DbReader.GetString(i);
                    switch (i)
                    {
                        case 0: auftragsposition.Vorgangsnummer = col; break;
                        case 1:
                            col = Regex.Replace(col, @"\s+", "");
                            auftragsposition.Artikelnummer = col;
                            break;
                        case 2:                            
                            auftragsposition.Gesamtpreis = col;
                            break;
                        case 3: auftragsposition.Bezeichnung = col; break;
                        case 4: auftragsposition.Anzahl = col; break;
                        case 5: auftragsposition.Mengeneinheit = col; break;
                    }
                }
                if (Int32.Parse(auftragsposition.Anzahl) > 0 && (auftragsposition.Artikelnummer.Length > 0) && String.Compare(auftragsposition.Artikelnummer, "0") != 0)
                {
                    try
                    {
                        Double.Parse(auftragsposition.Gesamtpreis);                        
                        List.Add(auftragsposition);
                    }catch(Exception e)
                    {
                        Console.WriteLine("Bedachter Fehler: "+e.ToString());
                    }
                }
            }

            DbCommand = database.CreateCommand();
            DbCommand.CommandText = "SELECT Vorgangsnummer, Stuecklistennummer, Gesamtpreis, Bezeichnung1, Menge, Mengeneinheit FROM " + StueckListenTabelle;
            DbReader = DbCommand.ExecuteReader();
            fCount = DbReader.FieldCount;
            //Stuecklistendaten durchgehen

            while (DbReader.Read())
            {
                Auftragsposition stuecklistenposition = new Auftragsposition();
                for (int i = 0; i < fCount; i++)
                {
                    String col = DbReader.GetString(i);
                    switch (i)
                    {
                        case 0: stuecklistenposition.Vorgangsnummer = col; break;
                        case 1:
                            col = Regex.Replace(col, @"\s+", "");
                            stuecklistenposition.Artikelnummer = "S" + col;
                            break;
                        case 2: stuecklistenposition.Gesamtpreis = col; break;
                        case 3: stuecklistenposition.Bezeichnung = col; break;
                        case 4: stuecklistenposition.Anzahl = col; break;
                        case 5: stuecklistenposition.Mengeneinheit = col; break;
                    }

                }
                if (Int32.Parse(stuecklistenposition.Anzahl) > 0 && (stuecklistenposition.Artikelnummer.Length > 0) && String.Compare(stuecklistenposition.Artikelnummer, "0") != 0)
                {
                    try
                    {
                        Double.Parse(stuecklistenposition.Gesamtpreis);                   
                        //Damit man aus der Liste weiß, ob es ein Stücklistenartikel ist oder nicht
                        stuecklistenposition.Stuecklistenartikel = true;
                        List.Add(stuecklistenposition);
                        // Console.WriteLine("Debug: " + stuecklistenposition.ToString());
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Achtung Fehler: " + e.ToString());
                    }
                }
                else
                {
//                    Console.WriteLine("Debug: " + stuecklistenposition.ToString() + " übersprungen.");
                }

            }
            try
            {
                database.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine("Fehler beim Schließen der Datenbank: " + e.ToString());
            }            


        }

    }
}
