using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aqualight.Artikelstamm
{
    [Serializable]
    class Artikel
    {

        public String Artikelnummer;
        public String Gesamtpreis;
        public String Bezeichnung;
        public String Anzahl;
        public String Mengeneinheit;
        public String Zoll;
        public double Gewicht;
        public String DimensionenKarton;
        public String DimensionenProdukt;
        public String Lieferzeit;        
        public bool Sperrgut;
        public bool Palette;
        public double EK;
        public double VK;
        public String[] Rabatt;
        public String Marke;
        public String EAN;
        public Artikel(){}
        public bool Aktiv;

        public String getArtikelnummer()
        {
            return Artikelnummer;
        }
    }
}
