using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aqualight.Artikelstamm
{
    class Pumpe : Artikel
    {
        public int Durchfluss;
        public int Foerderhoehe;
        public int Leistung;
        public bool Meerwassereinsatz;
        public bool Teicheinsatz;
        public bool Tauchbar;        
    }
}
