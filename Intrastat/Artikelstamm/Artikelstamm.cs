using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading.Tasks;

namespace Aqualight.Artikelstamm
{
    [Serializable]
    public class Artikelstamm
    {
        private HashSet<Artikel> hash;

        /// <summary>
        /// Holt den Artikelstamm aus dem Cache bzw. legt diesen erst an. Achtung: Braucht lange zum ausführen und viel Speicher
        /// am Besten sicherstellen, dass nur ein Objekt dieser Art exisitert.
        /// </summary>
        /// <returns>Hashset </returns>
        public Artikelstamm()
        {
            hash = new HashSet<Artikel>();

            try
            {
                IFormatter formatter = new BinaryFormatter();
                Stream stream = new FileStream("Artikelstamm.bin", FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

                if (stream.Length > 0)
                {
                    hash = (HashSet<Artikel>)formatter.Deserialize(stream);
                }
                stream.Close();
            }
            catch (FileNotFoundException fnfe)
            {                
                IFormatter formatter = new BinaryFormatter();
                Stream stream = new FileStream("Artikelstamm.bin", FileMode.Create, FileAccess.Write, FileShare.ReadWrite);
                hash = initializeArticles();
                formatter.Serialize(stream, hash);
                stream.Flush();
                stream.Close();
            }            
        }

        internal HashSet<Artikel> Hash
        {
            get
            {
                return hash;
            }

            set
            {
                hash = value;
            }
        }        

        /// <summary>
        /// Initializes the article hash if none was found or has to be updated
        /// </summary>
        private HashSet<Artikel> initializeArticles()
        {            
            OdbcConnection database = Database.Database.getDatabase();
            database.Open();
            OdbcCommand DbCommand = database.CreateCommand();
            DbCommand.CommandText = "SELECT Artikel, Hersteller, Mengeneinheit, Bezeichnung1  FROM artikel";
            OdbcDataReader DbReader = DbCommand.ExecuteReader();
            int fCount = DbReader.FieldCount;
            if (DbReader != null)
            {
                
                while (DbReader.Read())
                {
                    Artikel artikel = new Artikel();
                    for (int i = 0; i < fCount; i++)
                    {
                        String col = DbReader.GetString(i);
                        switch (i)
                        {
                            case 0: artikel.Artikelnummer = col; break;
                            case 1: artikel.Marke = col; break;
                            case 2: artikel.Mengeneinheit = col; break;
                            case 3: artikel.Bezeichnung = col; break;                            
                        }
                    }
                    hash.Add(artikel);
                }
            }
            enhanceArticlesFromExcel();
            return hash;
        }
        /// <summary>
        /// Artikelstamm muss schon initialisiert sein, dann werden alle Daten mit denen aus der Excelliste erweitert
        /// </summary>
        private void enhanceArticlesFromExcel()
        {
           List<BestellungElement> excel = ExcelParser.ExcelParser.getExcelParser();
           BestellungElement tmp = null;
           foreach (Artikel a in this.hash)
           {
                //tmp bezeichnet den passenden Artikel aus der Excelliste
                tmp = excel.Find(x => x.getArtikelnummer() == a.getArtikelnummer());
                if (tmp != null)
                {
                    a.DimensionenKarton = tmp.Verpackungsgroesse;
                    a.EAN = tmp.EAN;
                    a.Gewicht = tmp.Gewicht;
                    a.Palette = tmp.Palette;
                    a.Sperrgut = tmp.Sperrgut;
                    a.Aktiv = true;
                    a.Zoll = tmp.Zollnummer;
                }
                else
                {
                    //Das Produkt ist nicht in der Excelliste und wird deswegen inaktiv geschaltet.
                    a.Aktiv = false;
                }
           }
        }


    }
}

