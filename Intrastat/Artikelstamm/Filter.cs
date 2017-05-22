using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aqualight.Artikelstamm
{
    class Filter : Artikel
    {
        public String Aquariengroesse;
        public int Leistung; //W
        public int Durchfluss; //l/h
        public String Filtermedium;
        public bool Teichfilter;
        public String[] UVLeuchtmittel;

    }
}
