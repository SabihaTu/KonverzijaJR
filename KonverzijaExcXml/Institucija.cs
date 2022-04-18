using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KonverzijaExcXml
{
    class Institucija
    {
        public string jib, eMail, NazivInstitucije, eMailInstitucije;
        public int? vrstaNaknade, pravniOsnov, ucestalostIsplate;
        public string vrstaNaknadeStr, pravniOsnovStr, ucestalostIsplateStr;
        public DateTime? datumIsplate;
        public string datumIsplateStr;
    }
}
