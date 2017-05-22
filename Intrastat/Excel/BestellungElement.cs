using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aqualight
{
    [Serializable]
    class BestellungElement
    {
        public String Modelbezeichnung;
        public String EAN;
        public String Artikelnummer;
        public String Status;
        public Double Gewicht;
        public String Verpackungsgroesse;
        public String Zollnummer;
        public String Infos;
        public Double Preis;
        public String Rabatt;
        public int RabattStueckzahl;
        public Double PreisEndkunden;
        public bool Sperrgut;
        public bool Palette;


        public BestellungElement()
        {

        }
        public String getArtikelnummer()
        {
            return Artikelnummer;
        }
    
        public override String ToString() {
            return "Modelbezeichnung: " + Modelbezeichnung + " EAN: " + EAN + " Artikelnummer: " + Artikelnummer + " Status: " + Status + " Gewicht: " + Gewicht + " Verpackungsgroesse: " + Verpackungsgroesse + " Zollnummer: " + Zollnummer + " Infos: " + Infos + " Preis: " + Preis + " Rabatt: " + Rabatt + " RabattStueckzahl: " + RabattStueckzahl + " PreisEndkunden: " + PreisEndkunden + " ";
        }

    }
}
