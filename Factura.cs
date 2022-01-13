using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FacturasXMLAExcelManager
{
    internal class Factura
    {
        private String tipoDeDocumento;
        private String numeroDelDocumento;
        private String fechaDeDocumento;
        private String fechaContableDeDocumento;
        private String fechaDeVencimientoDeDocumento;
        private String codigoDeUnidadDeNegocio;
        private String rutCliente;
        private String direccionDelCliente;
        private String rutFacturador;
        private String codigoVendedor;
        private String codigoComisionista;
        private String probabilidad;
        private String listaPrecio;
        private String plazoPago;
        private String monedaDelDocumento;
        private String tasaDeCambio;
        private String montoAfecto;
        private String montoExento;
        private String montoIva;
        private String montoImpuestosEspecificos;
        private String montoIvaRetenido;
        private String montoImpuestosRetenidos;
        private String tipoDeDescuentoGlobal;
        private String descuentoGlobal;
        private String totalDelDocumento;
        private String deudaPendiente;
        private String tipoDocReferencia;
        private String numDocReferencia;
        private String fechaDocReferencia;
        private String codigoDelProducto;
        private String cantidad;
        private String unidad;
        private String precioUnitario;
        private String monedaDelDetalle;
        private String tasaDeCambio2;
        private String numeroDeSerie;
        private String numeroDeLote;
        private String fechaDeVencimiento;
        private String centroDeCostos;
        private String tipoDeDescuento;
        private String descuento;
        private String ubicacion;
        private String bodega;
        private String concepto1;
        private String concepto2;
        private String concepto3;
        private String concepto4;
        private String descripcion;
        private String descripcionAdicional;
        private String stock;
        private String Comentario1;
        private String Comentario2;
        private String Comentario3;
        private String Comentario4;
        private String Comentario5;
        private String codigoImpuestoEspecifico1;
        private String montoImpuestoEspecifico1;
        private String codigoImpuestoEspecifico2;
        private String montoImpuestoEspecifico2;
        private String modalidad;
        private String glosa;
        private String referencia;
        private String fechaDeComprometida;
        private String porcentajeCEEC;
        private String impuestoLey18211;
        private String ivaLey18211;
        private String codigoKitFlexible;
        private String ajusteIva;

       

        public Factura( )
        {
         
        }

        public Factura(string tipoDeDocumento, string numeroDelDocumento, string fechaDeDocumento, string fechaContableDeDocumento, string fechaDeVencimientoDeDocumento, string codigoDeUnidadDeNegocio, string rutCliente, string direccionDelCliente, string rutFacturador, string codigoVendedor, string codigoComisionista, string probabilidad, string listaPrecio, string plazoPago, string monedaDelDocumento, string tasaDeCambio, string montoAfecto, string montoExento, string montoIva, string montoImpuestosEspecificos, string montoIvaRetenido, string montoImpuestosRetenidos, string tipoDeDescuentoGlobal, string descuentoGlobal, string totalDelDocumento, string deudaPendiente, string tipoDocReferencia, string numDocReferencia, string fechaDocReferencia, string codigoDelProducto, string cantidad, string unidad, string precioUnitario, string monedaDelDetalle, String tasaDeCambio2, string numeroDeSerie, string numeroDeLote, string fechaDeVencimiento, string centroDeCostos, string tipoDeDescuento, string descuento, string ubicacion, string bodega, string concepto1, string concepto2, string concepto3, string concepto4, string descripcion, string descripcionAdicional, string stock, string comentario1, string comentario2, string comentario3, string comentario4, string comentario5, string codigoImpuestoEspecifico1, string montoImpuestoEspecifico1, string codigoImpuestoEspecifico2, string montoImpuestoEspecifico2, string modalidad, string glosa, string referencia, String fechaDeComprometida, string porcentajeCEEC, string impuestoLey18211, string ivaLey18211, string codigoKitFlexible, string ajusteIva)
        {
            this.TipoDeDocumento = tipoDeDocumento;
            this.NumeroDelDocumento = numeroDelDocumento;
            this.FechaDeDocumento = fechaDeDocumento;
            this.FechaContableDeDocumento = fechaContableDeDocumento;
            this.FechaDeVencimientoDeDocumento = fechaDeVencimientoDeDocumento;
            this.CodigoDeUnidadDeNegocio = codigoDeUnidadDeNegocio;
            this.RutCliente = rutCliente;
            this.DireccionDelCliente = direccionDelCliente;
            this.RutFacturador = rutFacturador;
            this.CodigoVendedor = codigoVendedor;
            this.CodigoComisionista = codigoComisionista;
            this.Probabilidad = probabilidad;
            this.ListaPrecio = listaPrecio;
            this.PlazoPago = plazoPago;
            this.MonedaDelDocumento = monedaDelDocumento;
            this.TasaDeCambio = tasaDeCambio;
            this.MontoAfecto = montoAfecto;
            this.MontoExento = montoExento;
            this.MontoIva = montoIva;
            this.MontoImpuestosEspecificos = montoImpuestosEspecificos;
            this.MontoIvaRetenido = montoIvaRetenido;
            this.MontoImpuestosRetenidos = montoImpuestosRetenidos;
            this.TipoDeDescuentoGlobal = tipoDeDescuentoGlobal;
            this.DescuentoGlobal = descuentoGlobal;
            this.TotalDelDocumento = totalDelDocumento;
            this.DeudaPendiente = deudaPendiente;
            this.TipoDocReferencia = tipoDocReferencia;
            this.NumDocReferencia = numDocReferencia;
            this.FechaDocReferencia = fechaDocReferencia;
            this.CodigoDelProducto = codigoDelProducto;
            this.Cantidad = cantidad;
            this.Unidad = unidad;
            this.PrecioUnitario = precioUnitario;
            this.MonedaDelDetalle = monedaDelDetalle;
            this.TasaDeCambio2 = tasaDeCambio2;
            this.NumeroDeSerie = numeroDeSerie;
            this.NumeroDeLote = numeroDeLote;
            this.FechaDeVencimiento = fechaDeVencimiento;
            this.CentroDeCostos = centroDeCostos;
            this.TipoDeDescuento = tipoDeDescuento;
            this.Descuento = descuento;
            this.Ubicacion = ubicacion;
            this.Bodega = bodega;
            this.Concepto1 = concepto1;
            this.Concepto2 = concepto2;
            this.Concepto3 = concepto3;
            this.Concepto4 = concepto4;
            this.Descripcion = descripcion;
            this.DescripcionAdicional = descripcionAdicional;
            this.Stock = stock;
            Comentario11 = comentario1;
            Comentario21 = comentario2;
            Comentario31 = comentario3;
            Comentario41 = comentario4;
            Comentario51 = comentario5;
            this.CodigoImpuestoEspecifico1 = codigoImpuestoEspecifico1;
            this.MontoImpuestoEspecifico1 = montoImpuestoEspecifico1;
            this.CodigoImpuestoEspecifico2 = codigoImpuestoEspecifico2;
            this.MontoImpuestoEspecifico2 = montoImpuestoEspecifico2;
            this.Modalidad = modalidad;
            this.Glosa = glosa;
            this.Referencia = referencia;
            this.FechaDeComprometida = fechaDeComprometida;
            this.PorcentajeCEEC = porcentajeCEEC;
            this.ImpuestoLey18211 = impuestoLey18211;
            this.IvaLey18211 = ivaLey18211;
            this.CodigoKitFlexible = codigoKitFlexible;
            this.AjusteIva = ajusteIva;
        }

        public string TipoDeDocumento { get => tipoDeDocumento; set => tipoDeDocumento = value; }
        public string NumeroDelDocumento { get => numeroDelDocumento; set => numeroDelDocumento = value; }
        public string FechaDeDocumento { get => fechaDeDocumento; set => fechaDeDocumento = value; }
        public string FechaContableDeDocumento { get => fechaContableDeDocumento; set => fechaContableDeDocumento = value; }
        public string FechaDeVencimientoDeDocumento { get => fechaDeVencimientoDeDocumento; set => fechaDeVencimientoDeDocumento = value; }
        public string CodigoDeUnidadDeNegocio { get => codigoDeUnidadDeNegocio; set => codigoDeUnidadDeNegocio = value; }
        public string RutCliente { get => rutCliente; set => rutCliente = value; }
        public string DireccionDelCliente { get => direccionDelCliente; set => direccionDelCliente = value; }
        public string RutFacturador { get => rutFacturador; set => rutFacturador = value; }
        public string CodigoVendedor { get => codigoVendedor; set => codigoVendedor = value; }
        public string CodigoComisionista { get => codigoComisionista; set => codigoComisionista = value; }
        public string Probabilidad { get => probabilidad; set => probabilidad = value; }
        public string ListaPrecio { get => listaPrecio; set => listaPrecio = value; }
        public string PlazoPago { get => plazoPago; set => plazoPago = value; }
        public string MonedaDelDocumento { get => monedaDelDocumento; set => monedaDelDocumento = value; }
        public string TasaDeCambio { get => tasaDeCambio; set => tasaDeCambio = value; }
        public string MontoAfecto { get => montoAfecto; set => montoAfecto = value; }
        public string MontoExento { get => montoExento; set => montoExento = value; }
        public string MontoIva { get => montoIva; set => montoIva = value; }
        public string MontoImpuestosEspecificos { get => montoImpuestosEspecificos; set => montoImpuestosEspecificos = value; }
        public string MontoIvaRetenido { get => montoIvaRetenido; set => montoIvaRetenido = value; }
        public string MontoImpuestosRetenidos { get => montoImpuestosRetenidos; set => montoImpuestosRetenidos = value; }
        public string TipoDeDescuentoGlobal { get => tipoDeDescuentoGlobal; set => tipoDeDescuentoGlobal = value; }
        public string DescuentoGlobal { get => descuentoGlobal; set => descuentoGlobal = value; }
        public string TotalDelDocumento { get => totalDelDocumento; set => totalDelDocumento = value; }
        public string DeudaPendiente { get => deudaPendiente; set => deudaPendiente = value; }
        public string TipoDocReferencia { get => tipoDocReferencia; set => tipoDocReferencia = value; }
        public string NumDocReferencia { get => numDocReferencia; set => numDocReferencia = value; }
        public string FechaDocReferencia { get => fechaDocReferencia; set => fechaDocReferencia = value; }
        public string CodigoDelProducto { get => codigoDelProducto; set => codigoDelProducto = value; }
        public string Cantidad { get => cantidad; set => cantidad = value; }
        public string Unidad { get => unidad; set => unidad = value; }
        public string PrecioUnitario { get => precioUnitario; set => precioUnitario = value; }
        public string MonedaDelDetalle { get => monedaDelDetalle; set => monedaDelDetalle = value; }
        public string TasaDeCambio2 { get => tasaDeCambio2; set => tasaDeCambio2 = value; }
        public string NumeroDeSerie { get => numeroDeSerie; set => numeroDeSerie = value; }
        public string NumeroDeLote { get => numeroDeLote; set => numeroDeLote = value; }
        public string FechaDeVencimiento { get => fechaDeVencimiento; set => fechaDeVencimiento = value; }
        public string CentroDeCostos { get => centroDeCostos; set => centroDeCostos = value; }
        public string TipoDeDescuento { get => tipoDeDescuento; set => tipoDeDescuento = value; }
        public string Descuento { get => descuento; set => descuento = value; }
        public string Ubicacion { get => ubicacion; set => ubicacion = value; }
        public string Bodega { get => bodega; set => bodega = value; }
        public string Concepto1 { get => concepto1; set => concepto1 = value; }
        public string Concepto2 { get => concepto2; set => concepto2 = value; }
        public string Concepto3 { get => concepto3; set => concepto3 = value; }
        public string Concepto4 { get => concepto4; set => concepto4 = value; }
        public string Descripcion { get => descripcion; set => descripcion = value; }
        public string DescripcionAdicional { get => descripcionAdicional; set => descripcionAdicional = value; }
        public string Stock { get => stock; set => stock = value; }
        public string Comentario11 { get => Comentario1; set => Comentario1 = value; }
        public string Comentario21 { get => Comentario2; set => Comentario2 = value; }
        public string Comentario31 { get => Comentario3; set => Comentario3 = value; }
        public string Comentario41 { get => Comentario4; set => Comentario4 = value; }
        public string Comentario51 { get => Comentario5; set => Comentario5 = value; }
        public string CodigoImpuestoEspecifico1 { get => codigoImpuestoEspecifico1; set => codigoImpuestoEspecifico1 = value; }
        public string MontoImpuestoEspecifico1 { get => montoImpuestoEspecifico1; set => montoImpuestoEspecifico1 = value; }
        public string CodigoImpuestoEspecifico2 { get => codigoImpuestoEspecifico2; set => codigoImpuestoEspecifico2 = value; }
        public string MontoImpuestoEspecifico2 { get => montoImpuestoEspecifico2; set => montoImpuestoEspecifico2 = value; }
        public string Modalidad { get => modalidad; set => modalidad = value; }
        public string Glosa { get => glosa; set => glosa = value; }
        public string Referencia { get => referencia; set => referencia = value; }
        public string FechaDeComprometida { get => fechaDeComprometida; set => fechaDeComprometida = value; }
        public string PorcentajeCEEC { get => porcentajeCEEC; set => porcentajeCEEC = value; }
        public string ImpuestoLey18211 { get => impuestoLey18211; set => impuestoLey18211 = value; }
        public string IvaLey18211 { get => ivaLey18211; set => ivaLey18211 = value; }
        public string CodigoKitFlexible { get => codigoKitFlexible; set => codigoKitFlexible = value; }
        public string AjusteIva { get => ajusteIva; set => ajusteIva = value; }
    }
}
