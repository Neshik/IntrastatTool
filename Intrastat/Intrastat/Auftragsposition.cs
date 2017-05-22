using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AquaLight.Intrastat
{
    
    [Serializable]
    class Auftragsposition
    {
        public String Vorgangsnummer;
        public String Artikelnummer;
        public String Gesamtpreis;
        public String Bezeichnung;
        public String Anzahl;
        public String Mengeneinheit;
        public String Zoll;
        public String Gewicht;
        public Boolean Stuecklistenartikel = false;

        public Auftragsposition()
        {

        }

        public Auftragsposition(String Vorgangsnummer,String Artikelnummer, String Gesamtpreis, String Bezeichnung, String Zoll, String Gewicht)
        {
            this.Vorgangsnummer = Vorgangsnummer;
            this.Artikelnummer = Artikelnummer;
            this.Gesamtpreis = Gesamtpreis;
            this.Bezeichnung = Bezeichnung;
            this.Zoll = Zoll;
        }

        public override String ToString(){
            return "Vorgangsnummer: " + Vorgangsnummer + " Artikelnummer " + Artikelnummer + " Gesamtpreis: " + Gesamtpreis + " Bezeichnung: " + Bezeichnung + " Anzahl: "+Anzahl;
        }
    }
}
