using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FacturasXMLAExcelManager
{
    internal class DireccionDeCCU
    {
        private String direccion;
        private String rut;
        private String folio;


        public DireccionDeCCU()
        {

        }

        public DireccionDeCCU(string direccion, string rut, string folio)
        {
            this.Direccion = direccion;
            this.Rut = rut;
            this.Folio = folio;
        }

        public string Direccion { get => direccion; set => direccion = value; }
        public string Rut { get => rut; set => rut = value; }
        public string Folio { get => folio; set => folio = value; }
    }
}
