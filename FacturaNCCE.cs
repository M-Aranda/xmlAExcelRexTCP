using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FacturasXMLAExcelManager
{
    internal class FacturaNCCE
    {
        //es documento de ciclo en Manager, para las facturas que son notas de crédito

        private String tipoDeDocumento; 
        private String numeroDelDocumento;
        private String fechaDeDocumento;
        private String fechaDeContableDeDocumento;
        private String fechaDeVencimientoDeDocumento;
        private String codigoUnidadDeNegocio;
        private String rutCliente;
        private String direccionCliente;
        private String rutFacturador;
        private String codigoVendedor; 
        private String codigoComisionista;
        private String probablidad;
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
        private String tipoDocumentoReferencia;
        private String numDocReferencia;
        private String fechaDocumentoDeReferencia;
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
        private String comentario1;
        private String comentario2;
        private String comentario3;
        private String comentario4;
        private String comentario5;
        private String codigoImpEspecial1;
        private String montoImpEspecial1;
        private String codigoImpEspecial2;
        private String montoImpEspecial2;
        private String modalidad;
        private String glosa;
        private String referencia;
        private String fechaDeComprometida;
        private String porcentajeCEEC;
        private String tipoDeDocumentoDeOrigen;
        private String numeroDocumentoDeOrigen;
        private String numeroDetalleOrigen;
        private String codigoKitFlexible;
        private String ajusteIva;

        public FacturaNCCE()
        {
           
        }

        public FacturaNCCE(string tipoDeDocumento, string numeroDelDocumento, string fechaDeDocumento, string fechaDeContableDeDocumento, string fechaDeVencimientoDeDocumento, string codigoUnidadDeNegocio, string rutCliente, string direccionCliente, string rutFacturador, string codigoVendedor, string codigoComisionista, string probablidad, string listaPrecio, string plazoPago, string monedaDelDocumento, string tasaDeCambio, string montoAfecto, string montoExento, string montoIva, string montoImpuestosEspecificos, string montoIvaRetenido, string montoImpuestosRetenidos, string tipoDeDescuentoGlobal, string descuentoGlobal, string totalDelDocumento, string deudaPendiente, string tipoDocumentoReferencia, string numDocReferencia, string fechaDocumentoDeReferencia, string codigoDelProducto, string cantidad, string unidad, string precioUnitario, string monedaDelDetalle, string tasaDeCambio2, string numeroDeSerie, string numeroDeLote, string fechaDeVencimiento, string centroDeCostos, string tipoDeDescuento, string descuento, string ubicacion, string bodega, string concepto1, string concepto2, string concepto3, string concepto4, string descripcion, string descripcionAdicional, string stock, string comentario1, string comentario2, string comentario3, string comentario4, string comentario5, string codigoImpEspecial1, string montoImpEspecial1, string codigoImpEspecial2, string montoImpEspecial2, string modalidad, string glosa, string referencia, string fechaDeComprometida, string porcentajeCEEC, string tipoDeDocumentoDeOrigen, string numeroDocumentoDeOrigen, string numeroDetalleOrigen, string codigoKitFlexible, string ajusteIva)
        {
            this.TipoDeDocumento = tipoDeDocumento;
            this.NumeroDelDocumento = numeroDelDocumento;
            this.FechaDeDocumento = fechaDeDocumento;
            this.FechaDeContableDeDocumento = fechaDeContableDeDocumento;
            this.FechaDeVencimientoDeDocumento = fechaDeVencimientoDeDocumento;
            this.CodigoUnidadDeNegocio = codigoUnidadDeNegocio;
            this.RutCliente = rutCliente;
            this.DireccionCliente = direccionCliente;
            this.RutFacturador = rutFacturador;
            this.CodigoVendedor = codigoVendedor;
            this.CodigoComisionista = codigoComisionista;
            this.Probablidad = probablidad;
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
            this.TipoDocumentoReferencia = tipoDocumentoReferencia;
            this.NumDocReferencia = numDocReferencia;
            this.FechaDocumentoDeReferencia = fechaDocumentoDeReferencia;
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
            this.Comentario1 = comentario1;
            this.Comentario2 = comentario2;
            this.Comentario3 = comentario3;
            this.Comentario4 = comentario4;
            this.Comentario5 = comentario5;
            this.CodigoImpEspecial1 = codigoImpEspecial1;
            this.MontoImpEspecial1 = montoImpEspecial1;
            this.CodigoImpEspecial2 = codigoImpEspecial2;
            this.MontoImpEspecial2 = montoImpEspecial2;
            this.Modalidad = modalidad;
            this.Glosa = glosa;
            this.Referencia = referencia;
            this.FechaDeComprometida = fechaDeComprometida;
            this.PorcentajeCEEC = porcentajeCEEC;
            this.TipoDeDocumentoDeOrigen = tipoDeDocumentoDeOrigen;
            this.NumeroDocumentoDeOrigen = numeroDocumentoDeOrigen;
            this.NumeroDetalleOrigen = numeroDetalleOrigen;
            this.CodigoKitFlexible = codigoKitFlexible;
            this.AjusteIva = ajusteIva;
        }

        public string TipoDeDocumento { get => tipoDeDocumento; set => tipoDeDocumento = value; }
        public string NumeroDelDocumento { get => numeroDelDocumento; set => numeroDelDocumento = value; }
        public string FechaDeDocumento { get => fechaDeDocumento; set => fechaDeDocumento = value; }
        public string FechaDeContableDeDocumento { get => fechaDeContableDeDocumento; set => fechaDeContableDeDocumento = value; }
        public string FechaDeVencimientoDeDocumento { get => fechaDeVencimientoDeDocumento; set => fechaDeVencimientoDeDocumento = value; }
        public string CodigoUnidadDeNegocio { get => codigoUnidadDeNegocio; set => codigoUnidadDeNegocio = value; }
        public string RutCliente { get => rutCliente; set => rutCliente = value; }
        public string DireccionCliente { get => direccionCliente; set => direccionCliente = value; }
        public string RutFacturador { get => rutFacturador; set => rutFacturador = value; }
        public string CodigoVendedor { get => codigoVendedor; set => codigoVendedor = value; }
        public string CodigoComisionista { get => codigoComisionista; set => codigoComisionista = value; }
        public string Probablidad { get => probablidad; set => probablidad = value; }
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
        public string TipoDocumentoReferencia { get => tipoDocumentoReferencia; set => tipoDocumentoReferencia = value; }
        public string NumDocReferencia { get => numDocReferencia; set => numDocReferencia = value; }
        public string FechaDocumentoDeReferencia { get => fechaDocumentoDeReferencia; set => fechaDocumentoDeReferencia = value; }
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
        public string Comentario1 { get => comentario1; set => comentario1 = value; }
        public string Comentario2 { get => comentario2; set => comentario2 = value; }
        public string Comentario3 { get => comentario3; set => comentario3 = value; }
        public string Comentario4 { get => comentario4; set => comentario4 = value; }
        public string Comentario5 { get => comentario5; set => comentario5 = value; }
        public string CodigoImpEspecial1 { get => codigoImpEspecial1; set => codigoImpEspecial1 = value; }
        public string MontoImpEspecial1 { get => montoImpEspecial1; set => montoImpEspecial1 = value; }
        public string CodigoImpEspecial2 { get => codigoImpEspecial2; set => codigoImpEspecial2 = value; }
        public string MontoImpEspecial2 { get => montoImpEspecial2; set => montoImpEspecial2 = value; }
        public string Modalidad { get => modalidad; set => modalidad = value; }
        public string Glosa { get => glosa; set => glosa = value; }
        public string Referencia { get => referencia; set => referencia = value; }
        public string FechaDeComprometida { get => fechaDeComprometida; set => fechaDeComprometida = value; }
        public string PorcentajeCEEC { get => porcentajeCEEC; set => porcentajeCEEC = value; }
        public string TipoDeDocumentoDeOrigen { get => tipoDeDocumentoDeOrigen; set => tipoDeDocumentoDeOrigen = value; }
        public string NumeroDocumentoDeOrigen { get => numeroDocumentoDeOrigen; set => numeroDocumentoDeOrigen = value; }
        public string NumeroDetalleOrigen { get => numeroDetalleOrigen; set => numeroDetalleOrigen = value; }
        public string CodigoKitFlexible { get => codigoKitFlexible; set => codigoKitFlexible = value; }
        public string AjusteIva { get => ajusteIva; set => ajusteIva = value; }
    }
}
