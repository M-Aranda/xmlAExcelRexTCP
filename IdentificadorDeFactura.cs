using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FacturasXMLAExcelManager
{
    internal class IdentificadorDeFactura
    {
        private String folio;
        private String rut;

        public IdentificadorDeFactura()
        {
        }

        public IdentificadorDeFactura(string folio, string rut)
        {
            this.Folio = folio;
            this.Rut = rut;
        }

        public string Folio { get => folio; set => folio = value; }
        public string Rut { get => rut; set => rut = value; }
    }
}
