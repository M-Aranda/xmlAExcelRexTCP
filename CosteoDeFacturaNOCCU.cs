using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FacturasXMLAExcelManager
{
    internal class CosteoDeFacturaNOCCU
    {
        private String folio;
        private String rut;
        private String afecto;
        private String centroDeCosto;

        public CosteoDeFacturaNOCCU()
        {
        }

        public CosteoDeFacturaNOCCU(string folio, string rut, string afecto, string centroDeCosto)
        {
            this.Folio = folio;
            this.Rut = rut;
            this.Afecto = afecto;
            this.CentroDeCosto = centroDeCosto;
        }

        public string Folio { get => folio; set => folio = value; }
        public string Rut { get => rut; set => rut = value; }
        public string Afecto { get => afecto; set => afecto = value; }
        public string CentroDeCosto { get => centroDeCosto; set => centroDeCosto = value; }
    }
}
