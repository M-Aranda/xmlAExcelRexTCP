using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FacturasXMLAExcelManager
{
    internal class CdgItem
    {

        private String tpoCodigo;
        private String vlrCodigo;

        public CdgItem(string tpoCodigo, string vlrCodigo)
        {
            this.TpoCodigo = tpoCodigo;
            this.VlrCodigo = vlrCodigo;
        }
        public CdgItem()
        {

        }

        public string TpoCodigo { get => tpoCodigo; set => tpoCodigo = value; }
        public string VlrCodigo { get => vlrCodigo; set => vlrCodigo = value; }







    }
}
