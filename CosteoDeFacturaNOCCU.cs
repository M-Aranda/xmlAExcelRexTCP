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
        private String glosa;//la glosa son las observaciones

        private String fechaDeDocumento;
        private String fechaContableDocumento;
        private String fechaDeVencimientoDeDocumento;

        private String exento;


        public CosteoDeFacturaNOCCU()
        {
        }

        public CosteoDeFacturaNOCCU(string folio, string rut, string afecto, string centroDeCosto, string montoIva, string ajusteIva, string codigoDelProducto, string glosa, string fechaDeDocumento, string fechaContableDocumento, string fechaDeVencimientoDeDocumento, string exento)
        {
            this.Folio = folio;
            this.Rut = rut;
            this.Afecto = afecto;
            this.CentroDeCosto = centroDeCosto;
            this.MontoIva = montoIva;
            this.AjusteIva = ajusteIva;
            this.CodigoDelProducto = codigoDelProducto;
            this.Glosa = glosa;
            this.FechaDeDocumento = fechaDeDocumento;
            this.FechaContableDocumento = fechaContableDocumento;
            this.FechaDeVencimientoDeDocumento = fechaDeVencimientoDeDocumento;
            this.Exento = exento;
        }

        public string Folio { get => folio; set => folio = value; }
        public string Rut { get => rut; set => rut = value; }
        public string Afecto { get => afecto; set => afecto = value; }
        public string CentroDeCosto { get => centroDeCosto; set => centroDeCosto = value; }
        public string MontoIva { get => montoIva; set => montoIva = value; }
        public string AjusteIva { get => ajusteIva; set => ajusteIva = value; }
        public string CodigoDelProducto { get => codigoDelProducto; set => codigoDelProducto = value; }
        public string Glosa { get => glosa; set => glosa = value; }
        public string FechaDeDocumento { get => fechaDeDocumento; set => fechaDeDocumento = value; }
        public string FechaContableDocumento { get => fechaContableDocumento; set => fechaContableDocumento = value; }
        public string FechaDeVencimientoDeDocumento { get => fechaDeVencimientoDeDocumento; set => fechaDeVencimientoDeDocumento = value; }
        public string Exento { get => exento; set => exento = value; }
    }
}
