using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FacturasXMLAExcelManager
{
    internal class RegistroDeCCU
    {

        private String rut;
        private String folio;
        private String centro;

        public RegistroDeCCU()
        {
        }

        public RegistroDeCCU(string rut, string folio, string centro)
        {
            this.Rut = rut;
            this.Folio = folio;
            this.Centro = centro;
        }

        public string Rut { get => rut; set => rut = value; }
        public string Folio { get => folio; set => folio = value; }
        public string Centro { get => centro; set => centro = value; }
    }
}
