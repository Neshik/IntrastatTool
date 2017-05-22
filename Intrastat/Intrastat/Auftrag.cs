using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AquaLight.Intrastat
{


    [Serializable]
    class Auftrag
    {
        public String auftragsnummer;
        public String Laendercode;
        public String date;
        public String bruttowert;
        public String gewicht;
        public List<Auftragsposition> auftragspositionenListe;
        public bool dataComplete = false;

        public Auftrag(String auftragsnummer, String Laendercode, String date, String wert, String gewicht)
        {
            this.auftragsnummer = auftragsnummer;
            this.Laendercode = Laendercode;
            this.date = date;
            this.bruttowert = wert;
            this.gewicht = gewicht;
        }

        public Auftrag()
        {

        }
        public override String ToString()
        {
            return "Auftragsnummer: " + auftragsnummer + " Laendercode: " + Laendercode + " date: " + date + " Gesamtwert: " + bruttowert + " Gewicht: " + gewicht;
        }
    }
}
