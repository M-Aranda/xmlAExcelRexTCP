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
        private String montoIva;
        private String ajusteIva;
        private String codigoDelProducto;

        public CosteoDeFacturaNOCCU()
        {
        }

        public CosteoDeFacturaNOCCU(string folio, string rut, string afecto, string centroDeCosto, string montoIva, string ajusteIva, string codigoDelProducto)
        {
            this.folio = folio;
            this.rut = rut;
            this.afecto = afecto;
            this.centroDeCosto = centroDeCosto;
            this.montoIva = montoIva;
            this.ajusteIva = ajusteIva;
            this.codigoDelProducto = codigoDelProducto;
        }

        public string Folio { get => folio; set => folio = value; }
        public string Rut { get => rut; set => rut = value; }
        public string Afecto { get => afecto; set => afecto = value; }
        public string CentroDeCosto { get => centroDeCosto; set => centroDeCosto = value; }
        public string MontoIva { get => montoIva; set => montoIva = value; }
        public string AjusteIva { get => ajusteIva; set => ajusteIva = value; }
        public string CodigoDelProducto { get => codigoDelProducto; set => codigoDelProducto = value; }
    }
}
