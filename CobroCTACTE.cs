using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FacturasXMLAExcelManager
{
    internal class CobroCTACTE
    {

        private String fecha;
        private String monto;
        private String ctoCosto;

        public CobroCTACTE()
        {

        }

        public CobroCTACTE(string fecha, string monto, string ctoCosto)
        {
            this.Fecha = fecha;
            this.Monto = monto;
            this.CtoCosto = ctoCosto;
        }

        public string Fecha { get => fecha; set => fecha = value; }
        public string Monto { get => monto; set => monto = value; }
        public string CtoCosto { get => ctoCosto; set => ctoCosto = value; }
    }
}
