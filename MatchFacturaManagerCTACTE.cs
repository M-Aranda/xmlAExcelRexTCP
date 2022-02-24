using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FacturasXMLAExcelManager
{
    internal class MatchFacturaManagerCTACTE
    {

        private String fecha;
        private String monto;
        private String centro;

        public MatchFacturaManagerCTACTE()
        {

        }

        public MatchFacturaManagerCTACTE(string fecha, string monto, string centro)
        {
            this.fecha = fecha;
            this.monto = monto;
            this.centro = centro;
        }

        public string Fecha { get => fecha; set => fecha = value; }
        public string Monto { get => monto; set => monto = value; }
        public string Centro { get => centro; set => centro = value; }
    }
}
