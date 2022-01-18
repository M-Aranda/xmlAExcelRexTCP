using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FacturasXMLAExcelManager
{
    internal class DetalleDeFactura
    {

        private String nroLinDet;
        private List<CdgItem> cdgosDeItem;
        private String nmbItem;
        private String qtyItem;
        private String unmdItem;
        private String prcItem;
        private String codImpAdic;
        private String montoItem;

        public string NroLinDet { get => nroLinDet; set => nroLinDet = value; }
        internal List<CdgItem> CdgosDeItem { get => cdgosDeItem; set => cdgosDeItem = value; }
        public string NmbItem { get => nmbItem; set => nmbItem = value; }
        public string QtyItem { get => qtyItem; set => qtyItem = value; }
        public string UnmdItem { get => unmdItem; set => unmdItem = value; }
        public string PrcItem { get => prcItem; set => prcItem = value; }
        public string CodImpAdic { get => codImpAdic; set => codImpAdic = value; }
        public string MontoItem { get => montoItem; set => montoItem = value; }

        public DetalleDeFactura(string nroLinDet, List<CdgItem> cdgosDeItem, string nmbItem, string qtyItem, string unmdItem, string prcItem, string codImpAdic, string montoItem)
        {
            this.NroLinDet = nroLinDet;
            this.CdgosDeItem = cdgosDeItem;
            this.NmbItem = nmbItem;
            this.QtyItem = qtyItem;
            this.UnmdItem = unmdItem;
            this.PrcItem = prcItem;
            this.CodImpAdic = codImpAdic;
            this.MontoItem = montoItem;
        }

        public DetalleDeFactura()
        {
      
        }


    }
}
