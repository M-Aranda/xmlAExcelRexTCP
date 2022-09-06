using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Windows.Storage;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace FacturasXMLAExcelManager
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //El excel a subir es el del formato de importación de documentos contables con detalle
            //estos botones son para las otras funciones




            //hay que generar los siguientes excel:
            //uno para los documentos de CCU que sean FACE o FCEE, esos van contabilizados de una (documento con detalles (contabilizado))
            //uno para los documentos no CCU y las notas de credito de cualquier cliente (documento con detalle)
            //otro para las guias de despacho



        }

        private void button1_Click(object sender, EventArgs e)
        {



            //excelAPartirDeXML
            List<Factura> facturasAIngresar = new List<Factura>();
            List<FacturaContabilizada> facturasAIngresarContabilizadas = new List<FacturaContabilizada>();
            List<FacturaNCCE> facturasNCCE = new List<FacturaNCCE>();
            List<GuiaDeDespacho> guiasDeDespacho = new List<GuiaDeDespacho>();
           

            int cantidadFACE = 0;
            int cantidadFCEE = 0;
            int cantidadNCCE = 0;
            int cantidadDeGuiasDeDespacho = 0;




            string sFileName = "";

            OpenFileDialog choofdlog = new OpenFileDialog();
            choofdlog.Filter = "All Files (*.*)|*.*";
            choofdlog.FilterIndex = 1;
            choofdlog.Multiselect = true;

            string[] arrAllFiles= new string[] {};


            while (true)
            {
                if (choofdlog.ShowDialog() == DialogResult.OK)
                {
                    sFileName = choofdlog.FileName;
                    arrAllFiles = choofdlog.FileNames; //used when Multiselect = true
                    break;
                }
                else
                {
                    MessageBox.Show("Debe seleccionar archivos XML");
                    System.Environment.Exit(0);
                }
            }

            Boolean variasFacturas = true;
   
            if (variasFacturas == false)//si es una sola factura, lo cual no debiese darse en el uso de este programa
            {
    
                String URLString = sFileName; 

                Factura f = new Factura();


                f.TipoDeDocumento = getValue("TipoDTE", sFileName);


                //fijarse con el SII
                f.TipoDeDocumento = determinarTipoDeDocumento(f.TipoDeDocumento);

                //Las fechas son en formato dd/mm/yyyy
                f.NumeroDelDocumento = getValue("Folio", sFileName);
                f.FechaDeDocumento = getValue("FchEmis", sFileName);
                f.FechaContableDeDocumento = getValue("FchEmis", sFileName);//que es la fecha de cancelacion?
                f.FechaDeVencimientoDeDocumento = getValue("FchVenc", sFileName);
                f.CodigoDeUnidadDeNegocio = "0"; //getValue("Folio", sFileName);
                f.RutCliente = getValue("RUTRecep", sFileName);// o sea nosotoros getValue("TipoDTE", sFileName);
                f.DireccionDelCliente = getValue("DirRecep", sFileName);
                f.RutFacturador = getValue("RUTEmisor", sFileName);
                f.CodigoVendedor = "";// getValue("TipoDTE", sFileName);
                f.CodigoComisionista = "";// getValue("Folio", sFileName);
                f.Probabilidad = "";// getValue("Folio", sFileName);
                f.ListaPrecio = "";//getValue("TipoDTE", sFileName);
                f.PlazoPago = "";//getValue("Folio", sFileName);
                f.MonedaDelDocumento = "CLP";//getValue("Folio", sFileName);
                f.TasaDeCambio = "";// getValue("TipoDTE", sFileName);
                f.MontoAfecto = getValue("MntNeto", sFileName);

                f.MontoExento = getValue("MontoImp", sFileName);
                //ojo con el monto exento


                //Hay que ingresar en mantenedores, importador de datos, datos a importar,
                //Documentos con detalles (Contabilizado)
                //hay 3 escenarios: si es factura exenta de iva, todo es 0, menos el campo de exento


                //si es factura electronica


                //MontoImp?
                //Si la factura es electronica no afecta, el MontoImp = MontoExento


                f.MontoIva = getValue("IVA", sFileName);
                f.MontoImpuestosEspecificos = "";//getValue("Folio", sFileName);
                f.MontoIvaRetenido = "";//getValue("Folio", sFileName);
                f.MontoImpuestosRetenidos = "";// getValue("TipoDTE", sFileName);
                f.TipoDeDescuentoGlobal = "";//getValue("Folio", sFileName);
                f.DescuentoGlobal = "";//getValue("Folio", sFileName);
                f.TotalDelDocumento = getValue("MntTotal", sFileName);
               
                f.DeudaPendiente = getValue("MntTotal", sFileName);// deuda pendiente siempre es igual al monto total
                f.TipoDocReferencia = "";//getValue("Folio", sFileName);
                f.NumDocReferencia = "";//getValue("Folio", sFileName);
                f.FechaDocReferencia = "";//getValue("Folio", sFileName);
                f.CodigoDelProducto = "410103";//getValue("TipoDTE", sFileName);
                //El codigo del producto va a depender de lo que sea el item
                f.Cantidad = "1"; getValue("Folio", sFileName);
                f.Unidad = "S.U.M"; //getValue("Folio", sFileName);
                f.PrecioUnitario = getValue("MntTotal", sFileName);
                f.MonedaDelDetalle = "CLP";//getValue("CLP", sFileName);
                f.TasaDeCambio2 = "1";//getValue("TipoDTE", sFileName);
                f.NumeroDeSerie = "";//getValue("Folio", sFileName);
                f.NumeroDeLote = "";//getValue("Folio", sFileName);
                f.FechaDeVencimiento = "";// getValue("Folio", sFileName);
                f.CentroDeCostos = getValue("DirOrigen", sFileName);//no es CmnaDest el dato que me da la dirección, es DirOrigen
                f.TipoDeDescuento = "";//getValue("TipoDTE", sFileName);
                f.Descuento = "";//getValue("Folio", sFileName);
                f.Ubicacion = "";//getValue("Folio", sFileName);
                f.Bodega = "";//getValue("TipoDTE", sFileName);
                f.Concepto1 = "";//getValue("Folio", sFileName);
                f.Concepto2 = "";//getValue("Folio", sFileName);
                f.Concepto3 = "";//getValue("TipoDTE", sFileName);
                f.Concepto4 = "";//getValue("Folio", sFileName);
                f.Descripcion = "";//getValue("Folio", sFileName);
                f.DescripcionAdicional = "";//getValue("Folio", sFileName);
                f.Stock = "0";//getValue("Folio", sFileName);
                f.Comentario11 = "";// getValue("TipoDTE", sFileName);
                f.Comentario21 = "";//getValue("Folio", sFileName);
                f.Comentario31 = "";//getValue("Folio", sFileName);
                f.Comentario41 = "";//getValue("Folio", sFileName);
                f.Comentario51 = "";//getValue("Folio", sFileName);
                f.CodigoImpuestoEspecifico1 = "";// getValue("TipoDTE", sFileName);
                f.MontoImpuestoEspecifico1 = "";// getValue("Folio", sFileName);
                f.CodigoImpuestoEspecifico2 = "";//getValue("Folio", sFileName);
                f.MontoImpuestoEspecifico2 = "";//getValue("Folio", sFileName);
                f.Modalidad = "";//getValue("Folio", sFileName);
                f.Glosa = "FALTANTE";//getValue("Folio", sFileName);
                f.Referencia = "";//getValue("Folio", sFileName);
                f.FechaDeComprometida = "";//getValue("Folio", sFileName);
                f.PorcentajeCEEC = "";//getValue("Folio", sFileName);
                f.ImpuestoLey18211 = "";//getValue("Folio", sFileName);
                f.IvaLey18211 = "";//getValue("Folio", sFileName);
                f.CodigoKitFlexible = "";//getValue("Folio", sFileName);
                f.AjusteIva = "";//getValue("Folio", sFileName);


                facturasAIngresar.Add(f);

            }

            else
            {

                foreach (var item in arrAllFiles)
                {

                   
                    // los ruts de CCU son:
                    // 91041000-8
                    //96989120-4
                    //99501760-1
                    //99554560-8
                    //99586280-8


                    //el rut de COPEC es:
                    //99520000-7, y se supone que es el único que tiene que tener
                    //tratamiento especial


                    sFileName = item;
                    // XmlTextReader reader = new XmlTextReader(URLString);

                    Factura f = new Factura();

                    f.MontoAfecto = "0";
                    f.MontoExento = "0";
                    f.MontoIva = "0";
                    f.TotalDelDocumento = "0";

                    //Hay que ver como leer los detalles

                    XmlTextReader textReader = new XmlTextReader(sFileName);
                    textReader.Read();



                    List<String> datos = new List<String>();
                    List<List<String>> datosDeDatos = new List<List<string>>();


                    f.TipoDeDocumento = getValue("TipoDTE", sFileName);
                    f.TipoDeDocumento = determinarTipoDeDocumento(f.TipoDeDocumento);
                    f.NumeroDelDocumento = getValue("Folio", sFileName);

                    //Las fechas son en formato dd/mm/yyyy

                    f.FechaDeDocumento = convertirAFechaValida3(getValue("FchEmis", sFileName));
                    f.FechaContableDeDocumento = f.FechaDeDocumento;
                    f.FechaDeVencimientoDeDocumento = convertirAFechaValida3(getValue("FchVenc", sFileName));//convertirAFechaValida(getValue("FchVenc", sFileName));// fecha de vencimiento debe ser igual o mayor a fecha de emision
                    f.FechaDeVencimientoDeDocumento = validarQueFechaVencimientoNoPrecedaAFechaDeEmision(f.FechaDeDocumento, f.FechaDeVencimientoDeDocumento);

                    if (f.RutCliente == "99520000-7")
                    {
                        f.FechaDeVencimientoDeDocumento = calcularFechaDeVencimientoDeCopec(f.FechaDeDocumento);
                    }

                    // DateTime now = DateTime.Now;
                    //f.FechaContableDeDocumento = Convert.ToString(now.Date);//"dia actual"convertirAFechaValidaDesdeTranstecnia

                    //f.FechaDeDocumento = f.FechaContableDeDocumento;
                    //f.FechaDeVencimientoDeDocumento = f.FechaContableDeDocumento;


                    f.CodigoDeUnidadDeNegocio = "1"; //getValue("Folio", sFileName);
                    f.RutCliente = getValue("RUTEmisor", sFileName);
                    f.DireccionDelCliente = "Casa Matriz"; //getValue("DirRecep", sFileName);
                    f.RutFacturador = "";//getValue("RUTEmisor", sFileName);//getValue("RUTRecep", sFileName);
                    f.CodigoVendedor = "";// getValue("TipoDTE", sFileName);
                    f.CodigoComisionista = "";// getValue("Folio", sFileName);
                    f.Probabilidad = "";// getValue("Folio", sFileName);
                    f.ListaPrecio = "";//getValue("TipoDTE", sFileName);
                    f.PlazoPago = "P01";//getValue("Folio", sFileName); codigo de plazo pago?
                    f.MonedaDelDocumento = "CLP";//getValue("Folio", sFileName);
                    f.TasaDeCambio = "";// getValue("TipoDTE", sFileName);
                    f.MontoAfecto = getValue("MntNeto", sFileName);



                    f.MontoExento = getValue("MontoImp", sFileName);



                    if (f.TipoDeDocumento == "FCEE")
                    {
                        f.MontoExento = getValue("MntExe", sFileName);

                    }

                    //MontoImp?
                    //Si la factura es electronica  afecta, el MontoImp = MontoExento
                    //Si la factura es electronica no afecta o exenta MntExe = MontoExento


                    f.MontoIva = getValue("IVA", sFileName);
                    f.MontoImpuestosEspecificos = "";//getValue("Folio", sFileName);
                    f.MontoIvaRetenido = "";//getValue("Folio", sFileName);
                    f.MontoImpuestosRetenidos = "";// getValue("TipoDTE", sFileName);
                    f.TipoDeDescuentoGlobal = "";//getValue("Folio", sFileName);
                    f.DescuentoGlobal = "";//getValue("Folio", sFileName);
                    f.TotalDelDocumento = getValue("MntTotal", sFileName);
                    f.DeudaPendiente = getValue("MntTotal", sFileName);
                    f.TipoDocReferencia = "";//getValue("Folio", sFileName);
                    f.NumDocReferencia = "";//getValue("Folio", sFileName);
                    f.FechaDocReferencia = "";//getValue("Folio", sFileName);

                    //El codigo del producto varia, puede ser uno de los siguientes:

                    /*
                     
                    420710
                    420724E

 

        110904	Impuesto Específico	Generico		
		310101	Fletes Interplantas	Servicio	
		310201	Fletes Emprendedores	Servicio	
		310202	Recuperacion de Gastos EMP	Servicio	
		310203	Arriendo Vehiculos	Servicio	
		310301	Fletes Tradicional	Servicio	
		310302	Recuperacion de Gastos	Servicio	
		410101	Contratista PSCP	Gasto	
		410102	Fletes Acarreo	Gasto	
		410103	Fletes de terceros	Gasto	
					
	    410104	Petroleo	Gasto		
		410105	Servicios de Terceros	Gasto	
		410105E	Transporte de Pasajeros	Gasto	
		410106	Arriendo Vehiculos	Gasto	
		410107	Arriendo Leasing	Gasto	
		410108	Fletes emprendedores	Gasto	
		410109	Cuota Arriendo Vehiculo (Socio)	Gasto	
		410110	Arriendo Vehiculo (Socio)	Gasto	
		420102	Honorarios	Gasto	
		420501	Contrato Mantenimiento	Gasto	
	    420502	Mantención Equipamiento (Bs.muebles	Gasto		
		420503	Mant. Inmuebles	Gasto	
		420504	Lubricantes	Gasto	
		420505	Insumos y Repuestos	Gasto	
		420506	Mantención Neumáticos	Gasto	
		420507	Otras Mantenciones	Gasto	
		420701	Arriendo de Oficinas	Gasto	
		420702	Gastos Comunes	Gasto	
		420703	Insumos de Aseo	Gasto	
		420704	Insumos de Oficinas	Gasto	
	    420705	Soporte Computacional	Gasto		
		420706	Energia Electrica	Gasto	
		420707	Agua	Gasto	
		420708	Gas	Gasto	
		420709	Telefonia y Comunicación	Gasto	
		420710	Gastos de Supermercados	Gasto	
		420711	Servicios de Correos	Gasto	
		420712	Servicios de Vigilancia	Gasto	
		420713	Gastos Notariales	Gasto	
		420714	Sanitizaciones	Gasto	
	    420715	Evaluaciones Medicas	Gasto		
		420716	Ropa de Trabajo del Personal	Gasto	
		420717	Arriendo de Maquinarias y Equipos	Gasto	
		420718	Seguros Camionetas, Camiones y otros	Gasto	
		420719	Seguros Oficinas	Gasto	
		420720	Patentes Municipales	Gasto	
		420721	Patentes y Permisos de Vehiculos	Gasto	
		420722	Revisiones Tecnicas	Gasto	
		420723	Infracciones de Transito	Gasto	
		420724	Faltantes	Gasto	
	    420724E	Faltantes exento	Gasto		
		420724ESP	Faltantes con impuesto específico	Gasto	
		420725	Gastos Varios	Gasto	
		420726	Honorarios de Auditoria Externa	Gasto	
		420727	Asesoria Legal	Gasto	
		420804	Gastos Bancarios	Gasto	

*/





                    f.CodigoDelProducto = "420724";//getValue("TipoDTE", sFileName);
                    f.Cantidad = "1"; //getValue("Folio", sFileName);
                    f.Unidad = "S.U.M"; //getValue("Folio", sFileName);
                    f.PrecioUnitario = getValue("MntNeto", sFileName);
                    f.MonedaDelDetalle = "CLP";
                    f.TasaDeCambio2 = "1";//getValue("TipoDTE", sFileName);
                    f.NumeroDeSerie = "";//getValue("Folio", sFileName);
                    f.NumeroDeLote = "";//getValue("Folio", sFileName);
                    f.FechaDeVencimiento = "";// getValue("Folio", sFileName);
                    f.CentroDeCostos = getValue("DirOrigen", sFileName);//no es CmnaDest el dato que me da la dirección, es DirOrigen



       

                    //determinar a donde se costea
                    //los codigos de centros de costo son (numero de la izquierda: TCP, numero de la derecha: PSCP): 

                    //203 / 303   Administracion
                    //204 / 304   Interplantas
                    //208 / 308   Emprendedores
                    //205 / 305   Illapel
                    //207 / 307   San Antonio
                    //200 / 300   Melipilla
                    //206 / 306   Santiago
                    //201 / 301   Rancagua
                    //202 / 302   Curico


                    //si el rut del receptor es 78462150-2, el costeo es para TCP, si
                    //es 78877610-1, es para PSCP
                    String rutDeReceptor= getValue("RutReceptor", sFileName);
                    Boolean esPSCP = false;
                    if (rutDeReceptor== "78877610-1")
                    {
                        esPSCP = true;
                    }

                    f.CentroDeCostos = determinarCentroDeCosto(f.CentroDeCostos, esPSCP);
                    f.CodigoDelProducto = "420724";
                

                    f.TipoDeDescuento = "";//getValue("TipoDTE", sFileName);
                    f.Descuento = "";//getValue("Folio", sFileName);
                    f.Ubicacion = "";//getValue("Folio", sFileName);
                    f.Bodega = "";//getValue("TipoDTE", sFileName);
                    f.Concepto1 = "";//getValue("Folio", sFileName);
                    f.Concepto2 = "";//getValue("Folio", sFileName);
                    f.Concepto3 = "";//getValue("TipoDTE", sFileName);
                    f.Concepto4 = "";//getValue("Folio", sFileName);
                    f.Descripcion = "";//getValue("Folio", sFileName);
                    f.DescripcionAdicional = "";//getValue("Folio", sFileName);
                    f.Stock = "0";//getValue("Folio", sFileName);
                    f.Comentario11 = "";// getValue("TipoDTE", sFileName);
                    f.Comentario21 = "";//getValue("Folio", sFileName);
                    f.Comentario31 = "";//getValue("Folio", sFileName);
                    f.Comentario41 = "";//getValue("Folio", sFileName);
                    f.Comentario51 = "";//getValue("Folio", sFileName);
                    f.CodigoImpuestoEspecifico1 = "";// getValue("TipoDTE", sFileName);
                    f.MontoImpuestoEspecifico1 = "";// getValue("Folio", sFileName);
                    f.CodigoImpuestoEspecifico2 = "";//getValue("Folio", sFileName);
                    f.MontoImpuestoEspecifico2 = "";//getValue("Folio", sFileName);
                    f.Modalidad = "";//getValue("Folio", sFileName);
                    if (f.TipoDeDocumento == "NCCE")
                    {
                        f.Modalidad = "3";
                    }

                    f.Glosa = "Factura de compra subida automaticamente";//getValue("Folio", sFileName);
                    f.Referencia = "";//getValue("Folio", sFileName);
                    f.FechaDeComprometida = "";//getValue("Folio", sFileName);
                    f.PorcentajeCEEC = "";//getValue("Folio", sFileName);
                    f.ImpuestoLey18211 = "";//getValue("Folio", sFileName);
                    f.IvaLey18211 = "";//getValue("Folio", sFileName);
                    f.CodigoKitFlexible = "";//getValue("Folio", sFileName);
                    f.AjusteIva = "";//getValue("Folio", sFileName);


                    //f.CodigoDelProducto = "420724";
                    f.PrecioUnitario = getValue("MntNeto", sFileName);
                    if (String.IsNullOrEmpty(f.MontoExento) == true)
                    {
                        f.MontoExento = "0";
                    }

                    String exentoAntes = f.MontoExento;
                    Boolean esFacturaDeEnvases = false;

                    //si es una factura de envase, el exento y el total debiesen ser el mismo
                    //if (f.NumeroDelDocumento == "128012803")
                    //{
                    //    MessageBox.Show(f.MontoAfecto);
                    //    MessageBox.Show(f.MontoExento);
                    //    MessageBox.Show(f.MontoIva);
                    //}
                    if (f.MontoAfecto == "0" && f.MontoExento == "0" && f.MontoIva == "0" && f.TotalDelDocumento != "0")
                    {
                        
                        f.CodigoDelProducto = "420724E";
                        f.PrecioUnitario = f.TotalDelDocumento;
                        f.MontoExento = f.TotalDelDocumento;
                        f.TipoDeDocumento = "FCEE";


                        esFacturaDeEnvases = true;
                    }


                    if (f.RutCliente != "91041000-8" & f.RutCliente != "96989120-4" & f.RutCliente != "99501760-1" & f.RutCliente != "99554560-8" & f.RutCliente != "99586280-8")
                    {

                        f.CodigoDelProducto = "110804";
                        if (f.RutCliente == "99520000-7")
                        {
                            f.CodigoDelProducto = "410104";
                        }
                        f.CentroDeCostos = "209";
                    }

                    if (f.TipoDeDocumento != "NCCE")
                    {

                        if (f.MontoIva == "0" && f.TipoDeDocumento=="FACE")
                        {
                            f.TipoDeDocumento = "FCEE";
                            f.CodigoDelProducto = "420724E";
                            
                        }



                        //si es una factura Exenta el iva y el afecto DEBEN ser 0
                        if (f.TipoDeDocumento == "FCEE")
                        {
                            f.MontoAfecto = "0";
                            f.MontoIva = "0";
                            f.PrecioUnitario = f.MontoExento;    
                        }


                        if (f.TipoDeDocumento != "FCEE")
                        {
                            f.MontoIva = calcularIvaComoManager(f.MontoAfecto, f.MontoIva, f);
                            f.MontoExento = determinarNuevoValorDeExentoAPartirDeMultiplesImpuestos(sFileName);

                        }

                        
                        recalcularTotales(f);

                        switch (f.TipoDeDocumento)
                        {
                            case "FACE":
                                cantidadFACE++;
                                break;
                            case "FCEE":
                                cantidadFCEE++;
                                break;
                            case "NCCE":
                                cantidadNCCE++;
                                break;
                            case "guia de despacho":
                                cantidadDeGuiasDeDespacho++;
                                break;
                            default:
                                break;
                        }

                        //Dejar dirección vacia si es uno de los clientes sin direccion
                        f.DireccionDelCliente = direccionSiEsQueEnManagerNoTiene(f.RutCliente);

                      if (f.TipoDeDocumento!= "guia de despacho")
                        {
                            if (f.RutCliente == "91041000-8" || f.RutCliente == "96989120-4" || f.RutCliente == "99501760-1" || f.RutCliente == "99554560-8" || f.RutCliente == "99586280-8")
                            {
                                //se sube a documento con detalle (contabilizado)
                                FacturaContabilizada fc = new FacturaContabilizada();
                                fc = fc.convertirFacturaAIngresarAFacturaContabilizada(f);
                                if (fc.CentroDeCostos=="204")
                                {
                                    fc.CodigoDeUnidadDeNegocio = "2";
                                }

                                facturasAIngresarContabilizadas.Add(fc);
                            }
                            else if (f.RutCliente != "91041000-8" && f.RutCliente != "96989120-4" && f.RutCliente != "99501760-1" && f.RutCliente != "99554560-8" && f.RutCliente != "99586280-8")
                            {
                                //si es una FACE o FCEE que NO sea de CCU O es una nota de credito de quien sea
                                //entonces se sube a documento con detalle (sin contabilizar)
                                if (f.TipoDeDocumento == "FCEE")
                                {
                                    f.CodigoDelProducto = "110804E";
                                }

                                //tratar facturas de Entel
                                if (f.RutCliente == "92580000-7")
                                {

                                    String direccionReceptor = getValue("DirRecep", sFileName);

                                    //Tenemos 4 cuentas/contratos de itnernet fija con entel:
                                    //322068 (Talca), 328985 (Rancagua CCU), 350011 (Renca) y 60167627 (Rancagua, oficina central)

                                    switch (direccionReceptor)
                                    {
                                        case "ANTONIO MACEO (CALLE) 2693": // ES RENCA
                                            f.CentroDeCostos = "206";
                                            f.CodigoDelProducto = "420709";
                                            f.Glosa = "Internet fija de Renca (Entel)";
                                            break;
                                        case "LONGITUDINAL SUR KM (CARR) 245": // ES TALCA
                                            f.CentroDeCostos = "202";
                                            f.CodigoDelProducto = "420709";
                                            f.Glosa = "Internet fija de Talca (Entel)";
                                            break;
                                        case "RUTA H-30 (CALLE) 2333":// ES RANCAGUA (CCU)
                                            f.CentroDeCostos = "201";
                                            f.CodigoDelProducto = "420709";
                                            f.Glosa = "Internet fija de Rancagua (CCU) (Entel)";
                                            break;
                                        case "CALLE CINCO SUR 85, RANCAGUA":// ES RANCAGUA (Central)
                                            f.CentroDeCostos = "203";
                                            f.CodigoDelProducto = "420709";
                                            f.Glosa = "Internet fija de oficina central en Rancagua (Entel)";
                                            break;
                                        default:
                                            //no es ninguna de las direcciones anteriores
                                            f.CentroDeCostos = "209";
                                            f.CodigoDelProducto = "420709";
                                            f.Glosa = "Factura de Entel";
                                            break;

                                    }
                                }

                                //tratar facturas de Movistar
                                if (f.RutCliente== "76124890-1")
                                {
                                    f.CentroDeCostos = "203";
                                    f.CodigoDelProducto = "420709";
                                    f.Glosa = "Factura de Movistar";

                                }

                                if(f.RutCliente== "92580000-7")
                                {
                                    FacturaContabilizada fc = new FacturaContabilizada();
                                    fc = fc.convertirFacturaAIngresarAFacturaContabilizada(f);
                                    facturasAIngresarContabilizadas.Add(fc);
                                }
                                else
                                {
                                    facturasAIngresar.Add(f);
                                }
                                

                            }

                        }
                        else
                        {
                            GuiaDeDespacho g = new GuiaDeDespacho();
                            g.TipoDeDocumento = f.TipoDeDocumento;
                            g.NumeroDelDocumento = f.NumeroDelDocumento;
                            g.FechaDeDocumento = f.FechaDeDocumento;
                            g.FechaContableDeDocumento = f.FechaContableDeDocumento;
                            g.FechaDeVencimientoDeDocumento = f.FechaDeVencimientoDeDocumento;
                            g.CodigoDeUnidadDeNegocio = f.CodigoDeUnidadDeNegocio;
                            g.RutCliente = f.RutCliente;
                            g.DireccionDelCliente = f.DireccionDelCliente;
                            g.RutFacturador = f.RutFacturador;
                            g.CodigoVendedor = f.CodigoVendedor;
                            g.CodigoComisionista = f.CodigoComisionista;
                            g.Probabilidad = f.Probabilidad;
                            g.ListaPrecio = f.ListaPrecio;
                            g.PlazoPago = f.PlazoPago;
                            g.MonedaDelDocumento = f.MonedaDelDocumento;
                            g.TasaDeCambio = f.TasaDeCambio;
                            g.MontoAfecto = f.MontoAfecto;
                            g.MontoExento = f.MontoExento;
                            g.MontoIva = f.MontoIva;
                            g.MontoImpuestosEspecificos = f.MontoImpuestosEspecificos;
                            g.MontoIvaRetenido = f.MontoIvaRetenido;
                            g.MontoImpuestosRetenidos = f.MontoImpuestosRetenidos;
                            g.TipoDeDescuentoGlobal = f.TipoDeDescuentoGlobal;
                            g.DescuentoGlobal = f.DescuentoGlobal;
                            g.TotalDelDocumento = f.TotalDelDocumento;
                            g.DeudaPendiente = f.DeudaPendiente;
                            g.TipoDocReferencia = "";//no se ingresa en documento contable con detalles
                            g.NumDocReferencia = "";//no se ingresa en documento contable con detalles
                            g.FechaDocReferencia = "";// no se ingresa en documento contable con detalles
                            g.CodigoDelProducto = f.CodigoDelProducto;
                            g.Cantidad = f.Cantidad;
                            g.Unidad = f.Unidad;
                            g.PrecioUnitario = f.PrecioUnitario;
                            g.MonedaDelDetalle = f.MonedaDelDetalle;
                            g.TasaDeCambio2 = f.TasaDeCambio2;
                            g.NumeroDeSerie = f.NumeroDeSerie;
                            g.NumeroDeLote = f.NumeroDeLote;
                            g.FechaDeVencimiento = f.FechaDeVencimiento;
                            g.CentroDeCostos = f.CentroDeCostos;
                            g.TipoDeDescuento = f.TipoDeDescuento;
                            g.Descuento = f.Descuento;
                            g.Ubicacion = f.Ubicacion;
                            g.Bodega = f.Bodega;
                            g.Concepto1 = f.Concepto1;
                            g.Concepto2 = f.Concepto2;
                            g.Concepto3 = f.Concepto3;
                            g.Concepto4 = f.Concepto4;
                            g.Descripcion = f.Descripcion;
                            g.DescripcionAdicional = f.DescripcionAdicional;
                            g.Stock = f.Stock;
                            g.Comentario11 = f.Comentario11;
                            g.Comentario21 = f.Comentario21;
                            g.Comentario31 = f.Comentario31;
                            g.Comentario41 = f.Comentario41;
                            g.Comentario51 = f.Comentario51;
                            g.CodigoImpuestoEspecifico1 = "";
                            g.MontoImpuestoEspecifico1 = "";
                            g.CodigoImpuestoEspecifico2 = "";
                            g.MontoImpuestoEspecifico2 = "";
                            g.Modalidad = f.Modalidad;
                            g.Glosa = "GUIA DE DESPACHO, NO SUBIR";
                            g.Referencia = f.Referencia;
                            g.FechaDeComprometida = f.FechaDeComprometida;
                            g.PorcentajeCEEC = f.PorcentajeCEEC;
                            g.ImpuestoLey18211 = "";
                            g.IvaLey18211 = "";
                            g.CodigoKitFlexible = f.CodigoKitFlexible;
                            g.AjusteIva = f.AjusteIva;

                             guiasDeDespacho.Add(g);
                        }

                        
                    }
                    else
                    {

                        cantidadNCCE++;

                        FacturaNCCE facNCCE = convertirFacturaANCCE(f);
                        facNCCE.TipoDeDocumentoDeOrigen = determinarTipoDeDocumento(getValue("TpoDocRef",sFileName));
                        facNCCE.NumeroDocumentoDeOrigen = getValue("FolioRef", sFileName);

           

                        facNCCE.FechaDeContableDeDocumento = facNCCE.FechaDeContableDeDocumento;
                        facNCCE.FechaDeDocumento = facNCCE.FechaDeContableDeDocumento;
                        facNCCE.FechaDeVencimientoDeDocumento = facNCCE.FechaDeContableDeDocumento;

                        if (facNCCE.RutCliente != "91041000-8" & facNCCE.RutCliente != "96989120-4" & facNCCE.RutCliente != "99501760-1" & facNCCE.RutCliente != "99554560-8" & facNCCE.RutCliente != "99586280-8")
                        {
                            facNCCE.CodigoDelProducto = "110804";
                            facNCCE.CentroDeCostos = "209";

                            if (facNCCE.RutCliente == "99520000-7")
                            {
                                facNCCE.CodigoDelProducto = "410104";
                            }
                        }

                        //420724E
                        //si nota de credito hace referencia a un folio que es 0, la factura tiene que subirse a documento contable
                        if (facNCCE.RutCliente == "91041000-8" || facNCCE.RutCliente == "96989120-4" || facNCCE.RutCliente == "99501760-1" || facNCCE.RutCliente == "99554560-8" || facNCCE.RutCliente == "99586280-8")
                        {
                            Factura factuarNCCEADocumentoContable = new Factura();
                            factuarNCCEADocumentoContable.TipoDeDocumento = facNCCE.TipoDeDocumento;
                            factuarNCCEADocumentoContable.NumeroDelDocumento = facNCCE.NumeroDelDocumento;

                            factuarNCCEADocumentoContable.FechaDeDocumento = facNCCE.FechaDeDocumento;
                            factuarNCCEADocumentoContable.FechaContableDeDocumento = facNCCE.FechaDeContableDeDocumento;
                            factuarNCCEADocumentoContable.FechaDeVencimientoDeDocumento = facNCCE.FechaDeVencimientoDeDocumento;

                            factuarNCCEADocumentoContable.CodigoDeUnidadDeNegocio = facNCCE.CodigoUnidadDeNegocio;
                            factuarNCCEADocumentoContable.RutCliente = facNCCE.RutCliente;
                            factuarNCCEADocumentoContable.DireccionDelCliente = facNCCE.DireccionCliente;
                            factuarNCCEADocumentoContable.RutFacturador = facNCCE.RutFacturador;
                            factuarNCCEADocumentoContable.CodigoVendedor = facNCCE.CodigoVendedor;
                            factuarNCCEADocumentoContable.CodigoComisionista = facNCCE.CodigoComisionista;
                            factuarNCCEADocumentoContable.Probabilidad = facNCCE.Probablidad;
                            factuarNCCEADocumentoContable.ListaPrecio = facNCCE.ListaPrecio;
                            factuarNCCEADocumentoContable.PlazoPago = facNCCE.PlazoPago;
                            factuarNCCEADocumentoContable.MonedaDelDocumento = facNCCE.MonedaDelDocumento;
                            factuarNCCEADocumentoContable.TasaDeCambio = facNCCE.TasaDeCambio;
                            factuarNCCEADocumentoContable.MontoAfecto = facNCCE.MontoAfecto;
                            factuarNCCEADocumentoContable.MontoExento = facNCCE.MontoExento;
                            factuarNCCEADocumentoContable.MontoIva = facNCCE.MontoIva;
                            factuarNCCEADocumentoContable.MontoImpuestosEspecificos = facNCCE.MontoImpuestosEspecificos;
                            factuarNCCEADocumentoContable.MontoIvaRetenido = facNCCE.MontoIvaRetenido;
                            factuarNCCEADocumentoContable.MontoImpuestosRetenidos = facNCCE.MontoImpuestosRetenidos;
                            factuarNCCEADocumentoContable.TipoDeDescuentoGlobal = facNCCE.TipoDeDescuentoGlobal;
                            factuarNCCEADocumentoContable.DescuentoGlobal = facNCCE.DescuentoGlobal;
                            factuarNCCEADocumentoContable.TotalDelDocumento = facNCCE.TotalDelDocumento;
                            factuarNCCEADocumentoContable.DeudaPendiente = facNCCE.DeudaPendiente;
                            factuarNCCEADocumentoContable.TipoDocReferencia = "";//no se ingresa en documento contable con detalles
                            factuarNCCEADocumentoContable.NumDocReferencia = "";//no se ingresa en documento contable con detalles
                            factuarNCCEADocumentoContable.FechaDocReferencia = "";// no se ingresa en documento contable con detalles
                            factuarNCCEADocumentoContable.CodigoDelProducto = facNCCE.CodigoDelProducto;
                            factuarNCCEADocumentoContable.Cantidad = facNCCE.Cantidad;
                            factuarNCCEADocumentoContable.Unidad = facNCCE.Unidad;
                            factuarNCCEADocumentoContable.PrecioUnitario = facNCCE.PrecioUnitario;
                            factuarNCCEADocumentoContable.MonedaDelDetalle = facNCCE.MonedaDelDetalle;
                            factuarNCCEADocumentoContable.TasaDeCambio2 = facNCCE.TasaDeCambio2;
                            factuarNCCEADocumentoContable.NumeroDeSerie = facNCCE.NumeroDeSerie;
                            factuarNCCEADocumentoContable.NumeroDeLote = facNCCE.NumeroDeLote;
                            factuarNCCEADocumentoContable.FechaDeVencimiento = facNCCE.FechaDeVencimiento;
                            factuarNCCEADocumentoContable.CentroDeCostos = facNCCE.CentroDeCostos;
                            factuarNCCEADocumentoContable.TipoDeDescuento = facNCCE.TipoDeDescuento;
                            factuarNCCEADocumentoContable.Descuento = facNCCE.Descuento;
                            factuarNCCEADocumentoContable.Ubicacion = facNCCE.Ubicacion;
                            factuarNCCEADocumentoContable.Bodega = facNCCE.Bodega;
                            factuarNCCEADocumentoContable.Concepto1 = facNCCE.Concepto1;
                            factuarNCCEADocumentoContable.Concepto2 = facNCCE.Concepto2;
                            factuarNCCEADocumentoContable.Concepto3 = facNCCE.Concepto3;
                            factuarNCCEADocumentoContable.Concepto4 = facNCCE.Concepto4;
                            factuarNCCEADocumentoContable.Descripcion = facNCCE.Descripcion;
                            factuarNCCEADocumentoContable.DescripcionAdicional = facNCCE.DescripcionAdicional;
                            factuarNCCEADocumentoContable.Stock = facNCCE.Stock;
                            factuarNCCEADocumentoContable.Comentario11 = facNCCE.Comentario1;
                            factuarNCCEADocumentoContable.Comentario21 = facNCCE.Comentario2;
                            factuarNCCEADocumentoContable.Comentario31 = facNCCE.Comentario3;
                            factuarNCCEADocumentoContable.Comentario41 = facNCCE.Comentario4;
                            factuarNCCEADocumentoContable.Comentario51 = facNCCE.Comentario5;
                            factuarNCCEADocumentoContable.CodigoImpuestoEspecifico1 = "";
                            factuarNCCEADocumentoContable.MontoImpuestoEspecifico1 = "";
                            factuarNCCEADocumentoContable.CodigoImpuestoEspecifico2 = "";
                            factuarNCCEADocumentoContable.MontoImpuestoEspecifico2 = "";
                            factuarNCCEADocumentoContable.Modalidad = facNCCE.Modalidad;
                            factuarNCCEADocumentoContable.Glosa = "NOTA DE CREDITO CON FOLIO "+ facNCCE.NumeroDocumentoDeOrigen;
                            factuarNCCEADocumentoContable.Referencia = facNCCE.Referencia;
                            factuarNCCEADocumentoContable.FechaDeComprometida = facNCCE.FechaDeComprometida;
                            factuarNCCEADocumentoContable.PorcentajeCEEC = facNCCE.PorcentajeCEEC;
                            factuarNCCEADocumentoContable.ImpuestoLey18211 = "";
                            factuarNCCEADocumentoContable.IvaLey18211 = "";
                            factuarNCCEADocumentoContable.CodigoKitFlexible = facNCCE.CodigoKitFlexible;
                            factuarNCCEADocumentoContable.AjusteIva = facNCCE.AjusteIva;


                            FacturaContabilizada fcNCCE1 = new FacturaContabilizada();

                            if (fcNCCE1.CentroDeCostos == "204")
                            {
                                fcNCCE1.CodigoDeUnidadDeNegocio = "2";
                            }
                            facturasAIngresarContabilizadas.Add(fcNCCE1.convertirFacturaAIngresarAFacturaContabilizada(factuarNCCEADocumentoContable));
                        }
                        else
                        {
                            Factura facturaNCCEaFactura1 = new Factura();
                            facturaNCCEaFactura1.TipoDeDocumento = facNCCE.TipoDeDocumento;
                            facturaNCCEaFactura1.NumeroDelDocumento = facNCCE.NumeroDelDocumento;
                            facturaNCCEaFactura1.FechaDeDocumento = facNCCE.FechaDeDocumento;
                            facturaNCCEaFactura1.FechaContableDeDocumento = facNCCE.FechaDeContableDeDocumento;
                            facturaNCCEaFactura1.FechaDeVencimientoDeDocumento = facNCCE.FechaDeVencimientoDeDocumento;
                            facturaNCCEaFactura1.CodigoDeUnidadDeNegocio = facNCCE.CodigoUnidadDeNegocio;
                            facturaNCCEaFactura1.RutCliente = facNCCE.RutCliente;
                            facturaNCCEaFactura1.DireccionDelCliente = facNCCE.DireccionCliente;
                            facturaNCCEaFactura1.RutFacturador = facNCCE.RutFacturador;
                            facturaNCCEaFactura1.CodigoVendedor = facNCCE.CodigoVendedor;
                            facturaNCCEaFactura1.CodigoComisionista = facNCCE.CodigoComisionista;
                            facturaNCCEaFactura1.Probabilidad = facNCCE.Probablidad;
                            facturaNCCEaFactura1.ListaPrecio = facNCCE.ListaPrecio;
                            facturaNCCEaFactura1.PlazoPago = facNCCE.PlazoPago;
                            facturaNCCEaFactura1.MonedaDelDocumento = facNCCE.MonedaDelDocumento;
                            facturaNCCEaFactura1.TasaDeCambio = facNCCE.TasaDeCambio;
                            facturaNCCEaFactura1.MontoAfecto = facNCCE.MontoAfecto;
                            facturaNCCEaFactura1.MontoExento = facNCCE.MontoExento;
                            facturaNCCEaFactura1.MontoIva = facNCCE.MontoIva;
                            facturaNCCEaFactura1.MontoImpuestosEspecificos = facNCCE.MontoImpuestosEspecificos;
                            facturaNCCEaFactura1.MontoIvaRetenido = facNCCE.MontoIvaRetenido;
                            facturaNCCEaFactura1.MontoImpuestosRetenidos = facNCCE.MontoImpuestosRetenidos;
                            facturaNCCEaFactura1.TipoDeDescuentoGlobal = facNCCE.TipoDeDescuentoGlobal;
                            facturaNCCEaFactura1.DescuentoGlobal = facNCCE.DescuentoGlobal;
                            facturaNCCEaFactura1.TotalDelDocumento = facNCCE.TotalDelDocumento;
                            facturaNCCEaFactura1.DeudaPendiente = facNCCE.DeudaPendiente;
                            facturaNCCEaFactura1.TipoDocReferencia = "";//no se ingresa en documento contable con detalles
                            facturaNCCEaFactura1.NumDocReferencia = "";//no se ingresa en documento contable con detalles
                            facturaNCCEaFactura1.FechaDocReferencia = "";// no se ingresa en documento contable con detalles
                            facturaNCCEaFactura1.CodigoDelProducto = facNCCE.CodigoDelProducto;
                            facturaNCCEaFactura1.Cantidad = facNCCE.Cantidad;
                            facturaNCCEaFactura1.Unidad = facNCCE.Unidad;
                            facturaNCCEaFactura1.PrecioUnitario = facNCCE.PrecioUnitario;
                            facturaNCCEaFactura1.MonedaDelDetalle = facNCCE.MonedaDelDetalle;
                            facturaNCCEaFactura1.TasaDeCambio2 = facNCCE.TasaDeCambio2;
                            facturaNCCEaFactura1.NumeroDeSerie = facNCCE.NumeroDeSerie;
                            facturaNCCEaFactura1.NumeroDeLote = facNCCE.NumeroDeLote;
                            facturaNCCEaFactura1.FechaDeVencimiento = facNCCE.FechaDeVencimiento;
                            facturaNCCEaFactura1.CentroDeCostos = facNCCE.CentroDeCostos;
                            facturaNCCEaFactura1.TipoDeDescuento = facNCCE.TipoDeDescuento;
                            facturaNCCEaFactura1.Descuento = facNCCE.Descuento;
                            facturaNCCEaFactura1.Ubicacion = facNCCE.Ubicacion;
                            facturaNCCEaFactura1.Bodega = facNCCE.Bodega;
                            facturaNCCEaFactura1.Concepto1 = facNCCE.Concepto1;
                            facturaNCCEaFactura1.Concepto2 = facNCCE.Concepto2;
                            facturaNCCEaFactura1.Concepto3 = facNCCE.Concepto3;
                            facturaNCCEaFactura1.Concepto4 = facNCCE.Concepto4;
                            facturaNCCEaFactura1.Descripcion = facNCCE.Descripcion;
                            facturaNCCEaFactura1.DescripcionAdicional = facNCCE.DescripcionAdicional;
                            facturaNCCEaFactura1.Stock = facNCCE.Stock;
                            facturaNCCEaFactura1.Comentario11 = facNCCE.Comentario1;
                            facturaNCCEaFactura1.Comentario21 = facNCCE.Comentario2;
                            facturaNCCEaFactura1.Comentario31 = facNCCE.Comentario3;
                            facturaNCCEaFactura1.Comentario41 = facNCCE.Comentario4;
                            facturaNCCEaFactura1.Comentario51 = facNCCE.Comentario5;
                            facturaNCCEaFactura1.CodigoImpuestoEspecifico1 = "";
                            facturaNCCEaFactura1.MontoImpuestoEspecifico1 = "";
                            facturaNCCEaFactura1.CodigoImpuestoEspecifico2 = "";
                            facturaNCCEaFactura1.MontoImpuestoEspecifico2 = "";
                            facturaNCCEaFactura1.Modalidad = facNCCE.Modalidad;
                            facturaNCCEaFactura1.Glosa = "NOTA DE CREDITO CON FOLIO " + facNCCE.NumeroDocumentoDeOrigen;
                            facturaNCCEaFactura1.Referencia = facNCCE.Referencia;
                            facturaNCCEaFactura1.FechaDeComprometida = facNCCE.FechaDeComprometida;
                            facturaNCCEaFactura1.PorcentajeCEEC = facNCCE.PorcentajeCEEC;
                            facturaNCCEaFactura1.ImpuestoLey18211 = "";
                            facturaNCCEaFactura1.IvaLey18211 = "";
                            facturaNCCEaFactura1.CodigoKitFlexible = facNCCE.CodigoKitFlexible;
                            facturaNCCEaFactura1.AjusteIva = facNCCE.AjusteIva;


                            facturasAIngresar.Add(facturaNCCEaFactura1);

                            facturasNCCE.Add(facNCCE);
                        }

                        
                        

                        // si el exento es distinto a 0, las notas de credito tienen que tener el mismo tratamiento con las face cuyo exento es superior a 0
                        if(facNCCE.MontoExento!="0")
                        {
                            FacturaNCCE facNCCE2 = new FacturaNCCE();
                            facNCCE2 = convertirFacturaANCCE(f);
                            facNCCE2.PrecioUnitario = facNCCE2.MontoExento;
                            facNCCE2.CodigoDelProducto = "420724E";

                            facNCCE2.TipoDeDocumentoDeOrigen = facNCCE.TipoDeDocumentoDeOrigen;
                            facNCCE2.NumeroDocumentoDeOrigen = facNCCE.NumeroDocumentoDeOrigen;

                            facNCCE2.FechaDeContableDeDocumento = facNCCE2.FechaDeContableDeDocumento;
                            facNCCE2.FechaDeDocumento = facNCCE2.FechaDeContableDeDocumento;
                            facNCCE2.FechaDeVencimientoDeDocumento = facNCCE2.FechaDeContableDeDocumento;

                            if (facNCCE2.RutCliente != "91041000-8" & facNCCE2.RutCliente != "96989120-4" & facNCCE2.RutCliente != "99501760-1" & facNCCE2.RutCliente != "99554560-8" & facNCCE2.RutCliente != "99586280-8")
                            {
                                facNCCE2.CodigoDelProducto = "110804";
                                facNCCE2.CentroDeCostos = "209";

                                if (facNCCE2.RutCliente == "99520000-7")
                                {
                                    facNCCE2.CodigoDelProducto = "410104";
                                }
                            }


                            //si nota de credito hace referencia a un folio que es 0, la factura tiene que subirse a documento contable
                            if (facNCCE2.RutCliente == "91041000-8" || facNCCE2.RutCliente == "96989120-4" || facNCCE2.RutCliente == "99501760-1" || facNCCE2.RutCliente == "99554560-8" || facNCCE2.RutCliente == "99586280-8")
                            {
                                Factura factuarNCCEADocumentoContable2 = new Factura();
                                factuarNCCEADocumentoContable2.TipoDeDocumento = facNCCE2.TipoDeDocumento;
                                factuarNCCEADocumentoContable2.NumeroDelDocumento = facNCCE2.NumeroDelDocumento;
                                factuarNCCEADocumentoContable2.FechaDeDocumento = facNCCE2.FechaDeDocumento;
                                factuarNCCEADocumentoContable2.FechaContableDeDocumento = facNCCE2.FechaDeContableDeDocumento;

                                factuarNCCEADocumentoContable2.FechaDeVencimientoDeDocumento = facNCCE2.FechaDeVencimientoDeDocumento;
                                factuarNCCEADocumentoContable2.CodigoDeUnidadDeNegocio = facNCCE2.CodigoUnidadDeNegocio;
                                factuarNCCEADocumentoContable2.RutCliente = facNCCE2.RutCliente;
                                factuarNCCEADocumentoContable2.DireccionDelCliente = facNCCE2.DireccionCliente;
                                factuarNCCEADocumentoContable2.RutFacturador = facNCCE2.RutFacturador;
                                factuarNCCEADocumentoContable2.CodigoVendedor = facNCCE2.CodigoVendedor;
                                factuarNCCEADocumentoContable2.CodigoComisionista = facNCCE2.CodigoComisionista;
                                factuarNCCEADocumentoContable2.Probabilidad = facNCCE2.Probablidad;
                                factuarNCCEADocumentoContable2.ListaPrecio = facNCCE2.ListaPrecio;
                                factuarNCCEADocumentoContable2.PlazoPago = facNCCE2.PlazoPago;
                                factuarNCCEADocumentoContable2.MonedaDelDocumento = facNCCE2.MonedaDelDocumento;
                                factuarNCCEADocumentoContable2.TasaDeCambio = facNCCE2.TasaDeCambio;
                                factuarNCCEADocumentoContable2.MontoAfecto = facNCCE2.MontoAfecto;
                                factuarNCCEADocumentoContable2.MontoExento = facNCCE2.MontoExento;
                                factuarNCCEADocumentoContable2.MontoIva = facNCCE2.MontoIva;
                                factuarNCCEADocumentoContable2.MontoImpuestosEspecificos = facNCCE2.MontoImpuestosEspecificos;
                                factuarNCCEADocumentoContable2.MontoIvaRetenido = facNCCE2.MontoIvaRetenido;
                                factuarNCCEADocumentoContable2.MontoImpuestosRetenidos = facNCCE2.MontoImpuestosRetenidos;
                                factuarNCCEADocumentoContable2.TipoDeDescuentoGlobal = facNCCE2.TipoDeDescuentoGlobal;
                                factuarNCCEADocumentoContable2.DescuentoGlobal = facNCCE2.DescuentoGlobal;
                                factuarNCCEADocumentoContable2.TotalDelDocumento = facNCCE2.TotalDelDocumento;
                                factuarNCCEADocumentoContable2.DeudaPendiente = facNCCE2.DeudaPendiente;
                                factuarNCCEADocumentoContable2.TipoDocReferencia = "";//no se ingresa en documento contable con detalles
                                factuarNCCEADocumentoContable2.NumDocReferencia = "";//no se ingresa en documento contable con detalles
                                factuarNCCEADocumentoContable2.FechaDocReferencia = "";// no se ingresa en documento contable con detalles
                                factuarNCCEADocumentoContable2.CodigoDelProducto = facNCCE2.CodigoDelProducto;
                                factuarNCCEADocumentoContable2.Cantidad = facNCCE2.Cantidad;
                                factuarNCCEADocumentoContable2.Unidad = facNCCE2.Unidad;
                                factuarNCCEADocumentoContable2.PrecioUnitario = facNCCE2.PrecioUnitario;
                                factuarNCCEADocumentoContable2.MonedaDelDetalle = facNCCE2.MonedaDelDetalle;
                                factuarNCCEADocumentoContable2.TasaDeCambio2 = facNCCE2.TasaDeCambio2;
                                factuarNCCEADocumentoContable2.NumeroDeSerie = facNCCE2.NumeroDeSerie;
                                factuarNCCEADocumentoContable2.NumeroDeLote = facNCCE2.NumeroDeLote;
                                factuarNCCEADocumentoContable2.FechaDeVencimiento = facNCCE2.FechaDeVencimiento;
                                factuarNCCEADocumentoContable2.CentroDeCostos = facNCCE2.CentroDeCostos;
                                factuarNCCEADocumentoContable2.TipoDeDescuento = facNCCE2.TipoDeDescuento;
                                factuarNCCEADocumentoContable2.Descuento = facNCCE2.Descuento;
                                factuarNCCEADocumentoContable2.Ubicacion = facNCCE2.Ubicacion;
                                factuarNCCEADocumentoContable2.Bodega = facNCCE2.Bodega;
                                factuarNCCEADocumentoContable2.Concepto1 = facNCCE2.Concepto1;
                                factuarNCCEADocumentoContable2.Concepto2 = facNCCE2.Concepto2;
                                factuarNCCEADocumentoContable2.Concepto3 = facNCCE2.Concepto3;
                                factuarNCCEADocumentoContable2.Concepto4 = facNCCE2.Concepto4;
                                factuarNCCEADocumentoContable2.Descripcion = facNCCE2.Descripcion;
                                factuarNCCEADocumentoContable2.DescripcionAdicional = facNCCE2.DescripcionAdicional;
                                factuarNCCEADocumentoContable2.Stock = facNCCE2.Stock;
                                factuarNCCEADocumentoContable2.Comentario11 = facNCCE2.Comentario1;
                                factuarNCCEADocumentoContable2.Comentario21 = facNCCE2.Comentario2;
                                factuarNCCEADocumentoContable2.Comentario31 = facNCCE2.Comentario3;
                                factuarNCCEADocumentoContable2.Comentario41 = facNCCE2.Comentario4;
                                factuarNCCEADocumentoContable2.Comentario51 = facNCCE2.Comentario5;
                                factuarNCCEADocumentoContable2.CodigoImpuestoEspecifico1 = "";
                                factuarNCCEADocumentoContable2.MontoImpuestoEspecifico1 = "";
                                factuarNCCEADocumentoContable2.CodigoImpuestoEspecifico2 = "";
                                factuarNCCEADocumentoContable2.MontoImpuestoEspecifico2 = "";
                                factuarNCCEADocumentoContable2.Modalidad = facNCCE2.Modalidad;
                                factuarNCCEADocumentoContable2.Glosa = "NOTA DE CREDITO CON FOLIO "+ facNCCE2.NumeroDocumentoDeOrigen;
                                factuarNCCEADocumentoContable2.Referencia = facNCCE2.Referencia;
                                factuarNCCEADocumentoContable2.FechaDeComprometida = facNCCE2.FechaDeComprometida;
                                factuarNCCEADocumentoContable2.PorcentajeCEEC = facNCCE2.PorcentajeCEEC;
                                factuarNCCEADocumentoContable2.ImpuestoLey18211 = "";
                                factuarNCCEADocumentoContable2.IvaLey18211 = "";
                                factuarNCCEADocumentoContable2.CodigoKitFlexible = facNCCE2.CodigoKitFlexible;
                                factuarNCCEADocumentoContable2.AjusteIva = facNCCE2.AjusteIva;

                                FacturaContabilizada fcNCCE2 = new FacturaContabilizada();
                                if (fcNCCE2.CentroDeCostos == "204")
                                {
                                    fcNCCE2.CodigoDeUnidadDeNegocio = "2";
                                }
                                facturasAIngresarContabilizadas.Add(fcNCCE2.convertirFacturaAIngresarAFacturaContabilizada(factuarNCCEADocumentoContable2));
                            }
                            else
                            {

                                Factura facturaNCCEaFactura2 = new Factura();

                                facturaNCCEaFactura2.TipoDeDocumento = facNCCE2.TipoDeDocumento;
                                facturaNCCEaFactura2.NumeroDelDocumento = facNCCE2.NumeroDelDocumento;
                                facturaNCCEaFactura2.FechaDeDocumento = facNCCE2.FechaDeDocumento;
                                facturaNCCEaFactura2.FechaContableDeDocumento = facNCCE2.FechaDeContableDeDocumento;
                                facturaNCCEaFactura2.FechaDeVencimientoDeDocumento = facNCCE2.FechaDeVencimientoDeDocumento;
                                facturaNCCEaFactura2.CodigoDeUnidadDeNegocio = facNCCE2.CodigoUnidadDeNegocio;
                                facturaNCCEaFactura2.RutCliente = facNCCE2.RutCliente;
                                facturaNCCEaFactura2.DireccionDelCliente = facNCCE2.DireccionCliente;
                                facturaNCCEaFactura2.RutFacturador = facNCCE2.RutFacturador;
                                facturaNCCEaFactura2.CodigoVendedor = facNCCE2.CodigoVendedor;
                                facturaNCCEaFactura2.CodigoComisionista = facNCCE2.CodigoComisionista;
                                facturaNCCEaFactura2.Probabilidad = facNCCE2.Probablidad;
                                facturaNCCEaFactura2.ListaPrecio = facNCCE2.ListaPrecio;
                                facturaNCCEaFactura2.PlazoPago = facNCCE2.PlazoPago;
                                facturaNCCEaFactura2.MonedaDelDocumento = facNCCE2.MonedaDelDocumento;
                                facturaNCCEaFactura2.TasaDeCambio = facNCCE2.TasaDeCambio;
                                facturaNCCEaFactura2.MontoAfecto = facNCCE2.MontoAfecto;
                                facturaNCCEaFactura2.MontoExento = facNCCE2.MontoExento;
                                facturaNCCEaFactura2.MontoIva = facNCCE2.MontoIva;
                                facturaNCCEaFactura2.MontoImpuestosEspecificos = facNCCE2.MontoImpuestosEspecificos;
                                facturaNCCEaFactura2.MontoIvaRetenido = facNCCE2.MontoIvaRetenido;
                                facturaNCCEaFactura2.MontoImpuestosRetenidos = facNCCE2.MontoImpuestosRetenidos;
                                facturaNCCEaFactura2.TipoDeDescuentoGlobal = facNCCE2.TipoDeDescuentoGlobal;
                                facturaNCCEaFactura2.DescuentoGlobal = facNCCE2.DescuentoGlobal;
                                facturaNCCEaFactura2.TotalDelDocumento = facNCCE2.TotalDelDocumento;
                                facturaNCCEaFactura2.DeudaPendiente = facNCCE2.DeudaPendiente;
                                facturaNCCEaFactura2.TipoDocReferencia = "";//no se ingresa en documento contable con detalles
                                facturaNCCEaFactura2.NumDocReferencia = "";//no se ingresa en documento contable con detalles
                                facturaNCCEaFactura2.FechaDocReferencia = "";// no se ingresa en documento contable con detalles
                                facturaNCCEaFactura2.CodigoDelProducto = facNCCE2.CodigoDelProducto;
                                facturaNCCEaFactura2.Cantidad = facNCCE2.Cantidad;
                                facturaNCCEaFactura2.Unidad = facNCCE2.Unidad;
                                facturaNCCEaFactura2.PrecioUnitario = facNCCE2.PrecioUnitario;
                                facturaNCCEaFactura2.MonedaDelDetalle = facNCCE2.MonedaDelDetalle;
                                facturaNCCEaFactura2.TasaDeCambio2 = facNCCE2.TasaDeCambio2;
                                facturaNCCEaFactura2.NumeroDeSerie = facNCCE2.NumeroDeSerie;
                                facturaNCCEaFactura2.NumeroDeLote = facNCCE2.NumeroDeLote;
                                facturaNCCEaFactura2.FechaDeVencimiento = facNCCE2.FechaDeVencimiento;
                                facturaNCCEaFactura2.CentroDeCostos = facNCCE2.CentroDeCostos;
                                facturaNCCEaFactura2.TipoDeDescuento = facNCCE2.TipoDeDescuento;
                                facturaNCCEaFactura2.Descuento = facNCCE2.Descuento;
                                facturaNCCEaFactura2.Ubicacion = facNCCE2.Ubicacion;
                                facturaNCCEaFactura2.Bodega = facNCCE2.Bodega;
                                facturaNCCEaFactura2.Concepto1 = facNCCE2.Concepto1;
                                facturaNCCEaFactura2.Concepto2 = facNCCE2.Concepto2;
                                facturaNCCEaFactura2.Concepto3 = facNCCE2.Concepto3;
                                facturaNCCEaFactura2.Concepto4 = facNCCE2.Concepto4;
                                facturaNCCEaFactura2.Descripcion = facNCCE2.Descripcion;
                                facturaNCCEaFactura2.DescripcionAdicional = facNCCE2.DescripcionAdicional;
                                facturaNCCEaFactura2.Stock = facNCCE2.Stock;
                                facturaNCCEaFactura2.Comentario11 = facNCCE2.Comentario1;
                                facturaNCCEaFactura2.Comentario21 = facNCCE2.Comentario2;
                                facturaNCCEaFactura2.Comentario31 = facNCCE2.Comentario3;
                                facturaNCCEaFactura2.Comentario41 = facNCCE2.Comentario4;
                                facturaNCCEaFactura2.Comentario51 = facNCCE2.Comentario5;
                                facturaNCCEaFactura2.CodigoImpuestoEspecifico1 = "";
                                facturaNCCEaFactura2.MontoImpuestoEspecifico1 = "";
                                facturaNCCEaFactura2.CodigoImpuestoEspecifico2 = "";
                                facturaNCCEaFactura2.MontoImpuestoEspecifico2 = "";
                                facturaNCCEaFactura2.Modalidad = facNCCE2.Modalidad;
                                facturaNCCEaFactura2.Glosa = "NOTA DE CREDITO CON FOLIO " + facNCCE2.NumeroDocumentoDeOrigen;
                                facturaNCCEaFactura2.Referencia = facNCCE2.Referencia;
                                facturaNCCEaFactura2.FechaDeComprometida = facNCCE2.FechaDeComprometida;
                                facturaNCCEaFactura2.PorcentajeCEEC = facNCCE2.PorcentajeCEEC;
                                facturaNCCEaFactura2.ImpuestoLey18211 = "";
                                facturaNCCEaFactura2.IvaLey18211 = "";
                                facturaNCCEaFactura2.CodigoKitFlexible = facNCCE2.CodigoKitFlexible;
                                facturaNCCEaFactura2.AjusteIva = facNCCE2.AjusteIva;


                                facturasAIngresar.Add(facturaNCCEaFactura2);

                                //el formato para las fechas si es nota de credito es distinto parece
                                facturasNCCE.Add(facNCCE2);


                            }

                        }

                    }
                                   
                    //f.MontoExento = exentoAntes;
                    
                    if (f.MontoExento !="0" && String.IsNullOrEmpty(f.MontoExento)==false && f.TipoDeDocumento == "FACE"  && esFacturaDeEnvases==false)
                    {
                        f = new Factura();
                        sFileName = item;

                        f.MontoAfecto = "0";
                        f.MontoExento = "0";
                        f.MontoIva = "0";
                        f.TotalDelDocumento = "0";

                        f.TipoDeDocumento = getValue("TipoDTE", sFileName);

                        f.TipoDeDocumento = determinarTipoDeDocumento(f.TipoDeDocumento);
                        
                        f.NumeroDelDocumento = getValue("Folio", sFileName);
                        f.FechaDeDocumento = convertirAFechaValida3(getValue("FchEmis", sFileName));
                        f.FechaContableDeDocumento = f.FechaDeDocumento;
                        f.FechaDeVencimientoDeDocumento = convertirAFechaValida3(getValue("FchVenc", sFileName));//convertirAFechaValida(getValue("FchVenc", sFileName));// fecha de vencimiento debe ser igual o mayor a fecha de emision
                        f.FechaDeVencimientoDeDocumento = validarQueFechaVencimientoNoPrecedaAFechaDeEmision(f.FechaDeDocumento,f.FechaDeVencimientoDeDocumento);

                        if (f.RutCliente == "99520000-7")
                        {
                            f.FechaDeVencimientoDeDocumento = calcularFechaDeVencimientoDeCopec(f.FechaDeDocumento);
                        }

                        //DateTime fechaActual = DateTime.Now;
                        //f.FechaContableDeDocumento = convertirAFechaValidaDesdeTranstecnia(Convert.ToString(now.Date));//"dia actual"

                        //f.FechaDeDocumento = f.FechaContableDeDocumento;
                        //f.FechaDeVencimientoDeDocumento = f.FechaContableDeDocumento;

                        f.CodigoDeUnidadDeNegocio = "1"; //getValue("Folio", sFileName);
                        f.RutCliente = getValue("RUTEmisor", sFileName);
                        f.DireccionDelCliente = "Casa Matriz"; //getValue("DirRecep", sFileName);
                        f.RutFacturador = "";//getValue("RUTEmisor", sFileName);//getValue("RUTRecep", sFileName);
                        f.CodigoVendedor = "";// getValue("TipoDTE", sFileName);
                        f.CodigoComisionista = "";// getValue("Folio", sFileName);
                        f.Probabilidad = "";// getValue("Folio", sFileName);
                        f.ListaPrecio = "";//getValue("TipoDTE", sFileName);
                        f.PlazoPago = "P01";//getValue("Folio", sFileName); codigo de plazo pago?
                        f.MonedaDelDocumento = "CLP";//getValue("Folio", sFileName);
                        f.TasaDeCambio = "";// getValue("TipoDTE", sFileName);
                        f.MontoAfecto = getValue("MntNeto", sFileName);


                        f.MontoExento = getValue("MontoImp", sFileName);

                        if (f.TipoDeDocumento == "FCEE")
                        {
                            f.MontoExento = getValue("MntExe", sFileName);
                        }

                        f.MontoIva = getValue("IVA", sFileName);
                        f.MontoImpuestosEspecificos = "";//getValue("Folio", sFileName);
                        f.MontoIvaRetenido = "";//getValue("Folio", sFileName);
                        f.MontoImpuestosRetenidos = "";// getValue("TipoDTE", sFileName);
                        f.TipoDeDescuentoGlobal = "";//getValue("Folio", sFileName);
                        f.DescuentoGlobal = "";//getValue("Folio", sFileName);
                        f.TotalDelDocumento = getValue("MntTotal", sFileName);
                        f.DeudaPendiente = getValue("MntTotal", sFileName);
                        f.TipoDocReferencia = "";//getValue("Folio", sFileName);
                        f.NumDocReferencia = "";//getValue("Folio", sFileName);
                        f.FechaDocReferencia = "";//getValue("Folio", sFileName);
                        f.CodigoDelProducto = "420724";//getValue("TipoDTE", sFileName);
                        f.Cantidad = "1"; //getValue("Folio", sFileName);
                        f.Unidad = "S.U.M"; //getValue("Folio", sFileName);
                        f.PrecioUnitario = getValue("MntNeto", sFileName);
                        f.MonedaDelDetalle = "CLP";
                        f.TasaDeCambio2 = "1";//getValue("TipoDTE", sFileName);
                        f.NumeroDeSerie = "";//getValue("Folio", sFileName);
                        f.NumeroDeLote = "";//getValue("Folio", sFileName);
                        f.FechaDeVencimiento = "";// getValue("Folio", sFileName);
                        f.CentroDeCostos = getValue("DirOrigen", sFileName); //Este es el centro de costos


                        //si el rut del receptor es 78462150-2, el costeo es para TCP, si
                        //es 78877610-1, es para PSCP
                         rutDeReceptor = getValue("RutReceptor", sFileName);
                         esPSCP = false;
                        if (rutDeReceptor == "78877610-1")
                        {
                            esPSCP = true;
                        }

                        f.CentroDeCostos = determinarCentroDeCosto(f.CentroDeCostos, esPSCP);

                        f.TipoDeDescuento = "";//getValue("TipoDTE", sFileName);
                        f.Descuento = "";//getValue("Folio", sFileName);
                        f.Ubicacion = "";//getValue("Folio", sFileName);
                        f.Bodega = "";//getValue("TipoDTE", sFileName);
                        f.Concepto1 = "";//getValue("Folio", sFileName);
                        f.Concepto2 = "";//getValue("Folio", sFileName);
                        f.Concepto3 = "";//getValue("TipoDTE", sFileName);
                        f.Concepto4 = "";//getValue("Folio", sFileName);
                        f.Descripcion = "";//getValue("Folio", sFileName);
                        f.DescripcionAdicional = "";//getValue("Folio", sFileName);
                        f.Stock = "0";//getValue("Folio", sFileName);
                        f.Comentario11 = "";// getValue("TipoDTE", sFileName);
                        f.Comentario21 = "";//getValue("Folio", sFileName);
                        f.Comentario31 = "";//getValue("Folio", sFileName);
                        f.Comentario41 = "";//getValue("Folio", sFileName);
                        f.Comentario51 = "";//getValue("Folio", sFileName);
                        f.CodigoImpuestoEspecifico1 = "";// getValue("TipoDTE", sFileName);
                        f.MontoImpuestoEspecifico1 = "";// getValue("Folio", sFileName);
                        f.CodigoImpuestoEspecifico2 = "";//getValue("Folio", sFileName);
                        f.MontoImpuestoEspecifico2 = "";//getValue("Folio", sFileName);
                        f.Modalidad = "";//getValue("Folio", sFileName);
                        f.Glosa = "Factura de compra subida automaticamente";//getValue("Folio", sFileName);
                        f.Referencia = "";//getValue("Folio", sFileName);
                        f.FechaDeComprometida = "";//getValue("Folio", sFileName);
                        f.PorcentajeCEEC = "";//getValue("Folio", sFileName);
                        f.ImpuestoLey18211 = "";
                        f.IvaLey18211 = "";
                        f.CodigoKitFlexible = "";
                        f.AjusteIva = "";

                        f.PrecioUnitario = f.MontoExento;
                        f.CodigoDelProducto = "420724E";


                        if (f.RutCliente != "91041000-8" & f.RutCliente != "96989120-4" & f.RutCliente != "99501760-1" & f.RutCliente != "99554560-8" & f.RutCliente != "99586280-8")
                        {
                            f.CodigoDelProducto = "110804E";
                            
                            if (f.RutCliente == "99520000-7")
                            {
                                f.CodigoDelProducto = "410104E"; 
                            }
                            f.CentroDeCostos = "209";
                        }


                        f.PrecioUnitario = f.MontoExento;
                        
                        //f.MontoExento = determinarNuevoValorDeExentoAPartirDeMultiplesImpuestos(sFileName);
                        f.MontoIva = calcularIvaComoManager(f.MontoAfecto,f.MontoIva,f);
                        recalcularTotales(f);

                        if (f.CentroDeCostos == "209" && f.RutCliente!= "99520000-7")
                        {
                            f.CodigoDelProducto = "410104E";
                        }

                        if (f.RutCliente == "91041000-8" || f.RutCliente == "96989120-4" || f.RutCliente == "99501760-1" || f.RutCliente == "99554560-8" || f.RutCliente == "99586280-8")
                        {
                            //se sube a documento con detalle (contabilizado)
                            FacturaContabilizada fc = new FacturaContabilizada();
                            fc = fc.convertirFacturaAIngresarAFacturaContabilizada(f);
                            if (fc.CentroDeCostos == "204")
                            {
                                //hay 4 unidades de negocio:
                                //Porteo (1)
                                //Acarreo (2)
                                //Emprendedores (3)
                                //Administración (4)

                                fc.CodigoDeUnidadDeNegocio = "2";
                            }

                            //agregado el 02/08/2022
                            fc = validarMontoExentoDeFacturaContabilizada(fc);


                            facturasAIngresarContabilizadas.Add(fc);
                        }
                        else if (f.RutCliente != "91041000-8" && f.RutCliente != "96989120-4" && f.RutCliente != "99501760-1" && f.RutCliente != "99554560-8" && f.RutCliente != "99586280-8")
                        {
                            //si es una FACE o FCEE que NO sea de CCU O es una nota de credito de quien sea
                            //entonces se sube a documento con detalle (sin contabilizar)

                            if (f.TipoDeDocumento == "FCEE")
                            {
                                f.CodigoDelProducto = "110804E";
                            }

                            //agregado el 02/08/2022
                            f = validarMontoExentoDeFacturaNoContabilizada(f);
                            facturasAIngresar.Add(f);

                        } 

                    }


                }


            }


            int totalDeFacturas = cantidadFACE+cantidadFCEE+cantidadNCCE+cantidadDeGuiasDeDespacho;

            MessageBox.Show("Se procesaron "+cantidadFACE.ToString()+" facturas afectas, "+cantidadFCEE.ToString()+" facturas exentas, "+cantidadNCCE.ToString()+" notas de crédito y "+cantidadDeGuiasDeDespacho.ToString()+" guías de despacho. El total fue de "+totalDeFacturas);

            //14-04-2022: se agrega ciclo para tratar facturas de Eccusa y Cervecera

            foreach (var item in facturasAIngresarContabilizadas)
            {
                //si factura a contabilizar es de cervecera o de eccusa, debe quedar como pendiente
                if ((item.CentroDeCostos == "206 ECCUSA") || (item.CentroDeCostos == "306 ECCUSA") || (item.CentroDeCostos == "208 CERVECERA") || (item.CentroDeCostos == "308 CERVECERA"))
                {
                    //agregar comentario
                    switch (item.CentroDeCostos)
                    {
                        case "206 ECCUSA":
                            item.Glosa = "Factura de Eccusa";
                            break;
                        case "306 ECCUSA":
                            item.Glosa = "Factura de Eccusa";
                            break;
                        case "208 CERVECERA":
                            item.Glosa = "Factura de Cervecera";
                            break;
                        case "308 CERVECERA":
                            item.Glosa = "Factura de Cervecera";
                            break;
                        default:
                            break;
                    }


                    item.CentroDeCostos = "209";


                    Factura f = new Factura(item);
                    facturasAIngresar.Add(f);

                    
                }
            }

            //segun lo anterior, si la factura de CCU era de CERVECERA o de Talca, esta debe quedar pendiente, por lo que
            //se quita de las contabilizadas
            facturasAIngresarContabilizadas.RemoveAll(item => item.CentroDeCostos == "209");

            //A veces vienen facturas con mas de un impuesto, que dan problemas porque 
            //la suma del afecto, exento e iva no calza con el monto total. Casi siempre eso es porque al exento le falta
            //considerar un impuesto. Asi que lo sigueinte debiese ser un fragmento de codigo que itere sobre
            //todas las facturas y revise esos valores.


            //28/07/2022, se agrega estructura para manejar facturas afectas y notas de crédito cuyo exento está incompleto

            String pathDeDescargas = getCarpetaDeDescargas();
            pathDeDescargas = pathDeDescargas + "" + @"\CCU (Documentos con detalle (contabilizado)).xlsx";
            var archivo = new FileInfo(pathDeDescargas);
            SaveExcelFileFacturasCCU(facturasAIngresarContabilizadas, archivo);

            pathDeDescargas = getCarpetaDeDescargas();
            pathDeDescargas = pathDeDescargas + "" + @"\NO CCU y CCU pendientes (documentos con detalle).xlsx";
            archivo = new FileInfo(pathDeDescargas);
            SaveExcelFile(facturasAIngresar, archivo);


            pathDeDescargas = getCarpetaDeDescargas() + "" + @"\Guias de despacho (NO SUBIR).xlsx";
            archivo = new FileInfo(pathDeDescargas);
            SaveExcelFileGuiasDeDespacho(guiasDeDespacho, archivo);


            MessageBox.Show("Archivo de facturas, archivo de notas de credito con folio y archivo de guias de despacho creados en carpeta de descargas!");


        }


    

        private static async Task SaveExcelFile(List<Factura> facturas, FileInfo file)
        {
            var package = new ExcelPackage(file);
            var ws = package.Workbook.Worksheets.Add("Facturas");

            var range = ws.Cells["A1"].LoadFromCollection(facturas, true);

            range.AutoFitColumns();

            await package.SaveAsync();
        }



        private static async Task SaveExcelFileGuiasDeDespacho(List<GuiaDeDespacho> guiasdeDespacho, FileInfo file)
        {
            var package = new ExcelPackage(file);
            var ws = package.Workbook.Worksheets.Add("Guias de despacho");

            var range = ws.Cells["A1"].LoadFromCollection(guiasdeDespacho, true);

            range.AutoFitColumns();

            await package.SaveAsync();
        }

        private static async Task SaveExcelFileFacturasCCU(List<FacturaContabilizada> facturasContabilizadas, FileInfo file)
        {
            var package = new ExcelPackage(file);
            var ws = package.Workbook.Worksheets.Add("Facturas contabilizadas");

            var range = ws.Cells["A1"].LoadFromCollection(facturasContabilizadas, true);

            range.AutoFitColumns();

            await package.SaveAsync();
        }


        private static async Task SaveExcelFileCosteoFacturasCCU(List<RegistroCruzadoConInformacionDeCCU> facturasDeCCUCosteadas, FileInfo file)
        {
            var package = new ExcelPackage(file);
            var ws = package.Workbook.Worksheets.Add("Facturas de CCU costeadas");

            var range = ws.Cells["A1"].LoadFromCollection(facturasDeCCUCosteadas, true);

            range.AutoFitColumns();

            await package.SaveAsync();
        }



        public string getValue(string _clave, String pathDelArchivo)
        {
            XmlTextReader textReader = new XmlTextReader(pathDelArchivo);
            textReader.Read();
            while (textReader.Read())
            {
                textReader.MoveToElement();
                if (textReader.Name == _clave)
                {
                    textReader.Read();
                    if (textReader.Value.ToString().Trim() != "")
                    {
                        string z = textReader.Value.ToString().Trim();
                        textReader.Close();
                        return z;
                    }
                }
            }
            textReader.Close();
            return "";
        }




        public String convertirAFechaValida(String fechaAConvertir)
        {
            String fechaValida = "";

            string[] datos = fechaAConvertir.Split('-');

            //manager pide dd/mm/yyyy, pero en la ultima prueba solo tomo mm/dd/yyyy

            //MessageBox.Show(datos[0]);//ano
            //MessageBox.Show(datos[1]);//mes
            //MessageBox.Show(datos[2]);//dia
            fechaValida = datos[1]+"/"+datos[2]+"/" + datos[0];

            // return fechaValida;
            //lo cambie porque está raro
            return "1/28/2022";

        }

        public String convertirAFechaValida2(String fechaAConvertir)
        {
            String fechaValida = "";

            string[] datos = fechaAConvertir.Split('/');

            //manager pide dd/mm/yyyy, pero en la ultima prueba solo tomo mm/dd/yyyy

            fechaValida = datos[1] + "/" + datos[0] + "/" + datos[2];

             return fechaValida;
      

        }


        public String convertirAFechaValidaDesdeTranstecnia(String fechaAConvertir)
        {
            String fechaValida = "";

            string[] fechaDeTranstecnia = fechaAConvertir.Split(' ');

            String soloLaFechaDeTranstecnia = fechaDeTranstecnia[0];

            string [] partes= soloLaFechaDeTranstecnia.Split('/');

            fechaValida = partes[1]+"/"+ partes[0] + "/" + partes[2] + "";


             return fechaValida;
        

        }





        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
  

        }

        private void button3_Click(object sender, EventArgs e)
        {
          
        }

        private void button4_Click(object sender, EventArgs e)
        {
           

            
        }


        private String determinarCentroDeCosto(String centroDeCostoComoString, Boolean esPSCP)
        {
            String centroDeCosto = "";


            switch (centroDeCostoComoString)
            {
                case "CURICO":
                    centroDeCosto = "202";
                    if (esPSCP)
                    {
                        centroDeCosto = "302";
                    }
                    break;
                case "TALCA":
                    centroDeCosto = "202";
                    if (esPSCP)
                    {
                        centroDeCosto = "302";
                    }
                    break;
                case "RANCAGUA":
                    centroDeCosto = "201";
                    if (esPSCP)
                    {
                        centroDeCosto = "302";
                    }
                    break;
                case "MELIPILLA":
                    centroDeCosto = "200";
                    if (esPSCP)
                    {
                        centroDeCosto = "300";
                    }
                    break;
                case "SAN ANTONIO":
                    centroDeCosto = "207";
                    if (esPSCP)
                    {
                        centroDeCosto = "307";
                    }
                    break;
                case "CURAUMA":
                    centroDeCosto = "207";
                    if (esPSCP)
                    {
                        centroDeCosto = "307";
                    }
                    break;
                case "SANTIAGO SUR":
                    centroDeCosto = "206";
                    if (esPSCP)
                    {
                        centroDeCosto = "306";
                    }
                    break;
                case "STGO PN1500":
                    //es Eccusa
                    centroDeCosto = "206 ECCUSA";
                    if (esPSCP)
                    {
                        centroDeCosto = "306 ECCUSA";
                    }
                    break;
                case "STGO PN8000":
                    //es Cervecera
                    centroDeCosto = "208 CERVECERA";
                    if (esPSCP)
                    {
                        centroDeCosto = "308 CERVECERA";
                    }
                    break;
                case "SANTIAGO":
                    centroDeCosto = "206";
                    if (esPSCP)
                    {
                        centroDeCosto = "306";
                    }
                    break;
                case "ILLAPEL":
                    centroDeCosto = "205";
                    if (esPSCP)
                    {
                        centroDeCosto = "305";
                    }
                    break;
                case "COQUIMBO":
                    centroDeCosto = "204";
                    if (esPSCP)
                    {
                        centroDeCosto = "304";
                    }
                    break;
                case "OVALLE":
                    centroDeCosto = "204";
                    if (esPSCP)
                    {
                        centroDeCosto = "304";
                    }
                    break;
                default:
                    //Interplantas, cuando no es ninguno de los anteriores
                    centroDeCosto = "204";//"204";
                    if (esPSCP)
                    {
                        centroDeCosto = "304";//"304";
                    }
                    break;

            }

            return centroDeCosto;

        }


        private String determinarTipoDeDocumento(String codigoDeDocumento)
        {
            String tipoDeFactura= "";

            switch (codigoDeDocumento)
            {
                case "33":
                    tipoDeFactura = "FACE";
                    break;
                case "34":
                    tipoDeFactura = "FCEE";
                    break;
                case "61":
                    tipoDeFactura = "NCCE";
                    break;
                case "52":
                    //es guia de despacho, así que no se ingresa a Manager
                    tipoDeFactura = "guia de despacho";
                    break;
                case "51":
                    tipoDeFactura = "NDCE";
                    break;
                case "56":
                    tipoDeFactura = "NDCE";
                    break;
                default:
                    tipoDeFactura = codigoDeDocumento;
                    break;
            }


            return tipoDeFactura;


        }


        private String getCarpetaDeDescargas()
        {
            string localfolder = ApplicationData.Current.LocalFolder.Path;
            var array = localfolder.Split('\\');
            var username = array[2];
            string downloads = @"C:\Users\" + username + @"\Downloads";
          

            return downloads;
        }

        private FacturaNCCE convertirFacturaANCCE(Factura f)
        {
            FacturaNCCE facNCCE = new FacturaNCCE();

            facNCCE.TipoDeDocumento = f.TipoDeDocumento;
            facNCCE.NumeroDelDocumento = f.NumeroDelDocumento;
            facNCCE.FechaDeDocumento = f.FechaDeDocumento;
            facNCCE.FechaDeContableDeDocumento = f.FechaContableDeDocumento;
            facNCCE.FechaDeVencimientoDeDocumento = f.FechaDeVencimientoDeDocumento;
            facNCCE.CodigoUnidadDeNegocio = f.CodigoDeUnidadDeNegocio;
            facNCCE.RutCliente = f.RutCliente;
            facNCCE.DireccionCliente = f.DireccionDelCliente;
            facNCCE.RutFacturador = f.RutFacturador;
            facNCCE.CodigoVendedor = f.CodigoVendedor;
            facNCCE.CodigoComisionista = f.CodigoComisionista;
            facNCCE.Probablidad = f.Probabilidad;
            facNCCE.ListaPrecio = f.ListaPrecio;
            facNCCE.PlazoPago = f.PlazoPago;
            facNCCE.MonedaDelDocumento = f.MonedaDelDocumento;
            facNCCE.TasaDeCambio = f.TasaDeCambio;
            facNCCE.MontoAfecto = f.MontoAfecto;
            facNCCE.MontoExento = f.MontoExento;
            facNCCE.MontoIva = f.MontoIva;
            facNCCE.MontoImpuestosEspecificos = f.MontoImpuestosEspecificos;
            facNCCE.MontoIvaRetenido = f.MontoIvaRetenido;
            facNCCE.MontoImpuestosRetenidos = f.MontoImpuestosRetenidos;
            facNCCE.TipoDeDescuentoGlobal = f.TipoDeDescuentoGlobal;
            facNCCE.DescuentoGlobal = f.DescuentoGlobal;
            facNCCE.TotalDelDocumento = f.TotalDelDocumento;
            facNCCE.DeudaPendiente = f.DeudaPendiente;
            facNCCE.TipoDocumentoReferencia = f.TipoDocReferencia;
            facNCCE.NumDocReferencia = f.NumDocReferencia;
            facNCCE.FechaDocumentoDeReferencia = f.FechaDocReferencia;
            facNCCE.CodigoDelProducto = f.CodigoDelProducto;
            facNCCE.Cantidad = f.Cantidad;
            facNCCE.Unidad = f.Unidad;
            facNCCE.PrecioUnitario = f.PrecioUnitario;
            facNCCE.MonedaDelDetalle = f.MonedaDelDetalle;
            facNCCE.TasaDeCambio2 = f.TasaDeCambio2;
            facNCCE.NumeroDeSerie = f.NumeroDeSerie;
            facNCCE.NumeroDeLote = f.NumeroDeLote;
            facNCCE.FechaDeVencimiento = f.FechaDeVencimiento;
            facNCCE.CentroDeCostos = f.CentroDeCostos;
            facNCCE.TipoDeDescuento = f.TipoDeDescuento;
            facNCCE.Descuento = f.Descuento;
            facNCCE.Ubicacion = f.Ubicacion;
            facNCCE.Bodega = f.Bodega;
            facNCCE.Concepto1 = f.Concepto1;
            facNCCE.Concepto2 = f.Concepto2;
            facNCCE.Concepto3 = f.Concepto3;
            facNCCE.Concepto4 = f.Concepto4;
            facNCCE.Descripcion = f.Descripcion;
            facNCCE.DescripcionAdicional = f.DescripcionAdicional;
            facNCCE.Stock = f.Stock;
            facNCCE.Comentario1 = f.Comentario11;
            facNCCE.Comentario2 = f.Comentario21;
            facNCCE.Comentario3 = f.Comentario31;
            facNCCE.Comentario4 = f.Comentario41;
            facNCCE.Comentario5 = f.Comentario51;
            facNCCE.CodigoImpEspecial1 = f.CodigoImpuestoEspecifico1;
            facNCCE.MontoImpEspecial1 = f.MontoImpuestoEspecifico1;
            facNCCE.CodigoImpEspecial2 = f.CodigoImpuestoEspecifico2;
            facNCCE.MontoImpEspecial2 = f.MontoImpuestoEspecifico2;
            facNCCE.Modalidad = f.Modalidad;
            facNCCE.Glosa = f.Glosa;
            facNCCE.Referencia = f.Referencia;
            facNCCE.FechaDeComprometida = f.FechaDeComprometida;
            facNCCE.PorcentajeCEEC = "";
            facNCCE.TipoDeDocumentoDeOrigen = "";
            facNCCE.NumeroDocumentoDeOrigen = "";
            facNCCE.NumeroDetalleOrigen = "";
            facNCCE.CodigoKitFlexible = f.CodigoKitFlexible;
            facNCCE.AjusteIva = f.AjusteIva;

            return facNCCE;
        }

        private void button5_Click(object sender, EventArgs e)
        {

            

        }


     

        private String calcularIvaComoManager(String afecto, String iva, Factura f)
        {
            String valorDeIvaARetornar = iva;

            if (String.IsNullOrEmpty(afecto))
            {
                afecto = "0";
            }
            

            Double valorIvaCalculado = Math.Round((int.Parse(afecto) * 0.19));


            String valorIvaCalculadoComoString = valorIvaCalculado.ToString();

            if (valorIvaCalculadoComoString != iva)
            {
                valorDeIvaARetornar = valorIvaCalculadoComoString;
                //validar que si termina en .5 se redondee hacia arriba

                if (valorIvaCalculado % 0.5f == 0){
                    valorIvaCalculado = valorIvaCalculado + 1.0;
                                             
                }

            }
            
            return valorDeIvaARetornar;

        }

    
        private String determinarNuevoValorDeExentoAPartirDeMultiplesImpuestos(String sFileName)
        {
            XmlTextReader reader = new XmlTextReader(sFileName);
            Boolean esMontoDeImpuesto = false;
            String impuestosSumados = "";
            int valorDeImpuesto = 0;

            try
            {
                while (reader.Read())
                {
                    switch (reader.NodeType)
                    {
                        case XmlNodeType.Element: // The node is an element.
                            if (reader.Name == "MontoImp")
                            {
                                esMontoDeImpuesto = true;
                                Console.Write("<" + reader.Name);
                                while (reader.MoveToNextAttribute()) // Read the attributes.
                                    Console.Write(" " + reader.Name + "='" + reader.Value + "'");
                                Console.Write(">");
                            }


                            break;
                        case XmlNodeType.Text: //Display the text in each element.
                            if (esMontoDeImpuesto == true)
                            {
                                Console.WriteLine(reader.Value);
                                if (String.IsNullOrEmpty(reader.Value) == true)
                                {
                                    valorDeImpuesto = valorDeImpuesto + 0;
                                }
                                else
                                {
                                    valorDeImpuesto = valorDeImpuesto + int.Parse(reader.Value);
                                }

                            }

                            esMontoDeImpuesto = false;
                            break;

                    }
                }
            }
            catch (Exception e)
            {
                Console.Write(e);
                
                throw;
            }


            impuestosSumados = valorDeImpuesto.ToString();
            return impuestosSumados;
        }

        private void recalcularTotales(Factura f)
        {

            int afecto = int.Parse(f.MontoAfecto);
            int exento = int.Parse(f.MontoExento);
            int iva = int.Parse(f.MontoIva);
           

            int total = afecto + exento + iva;
            f.TotalDelDocumento=total.ToString();
            f.DeudaPendiente=total.ToString();
           
        }


 

        public String convertirAFechaValida3(String fechaAConvertir)
        {
            String fechaValida = "";
            if (String.IsNullOrEmpty(fechaAConvertir)!=true)
            {
                string[] datos = fechaAConvertir.Split('-');
                //MessageBox.Show(datos[0]);//ano
                //MessageBox.Show(datos[1]);//mes
                //MessageBox.Show(datos[2]);//dia
                fechaValida = datos[2] + "/" + datos[1] + "/" + datos[0];
            }
            else
            {
                //si la fecha de vencimiento no esta presente, entonces la fecha de vencimiento debiese ser igual a la fecha de emision
                fechaValida = "Fecha de vencimiento ausente en factura";
            }
            return fechaValida;

        }

        public String validarQueFechaVencimientoNoPrecedaAFechaDeEmision(String fechaDeEmision, String fechaDeVencimiento)
        {
            String fechaCorrectaDeVencimiento = fechaDeVencimiento;

            string[] datos = fechaDeEmision.Split('/');
            if (fechaDeVencimiento != "Fecha de vencimiento ausente en factura")
            {
                //fecha de vencimiento no esta vacia
                string[] datos2 = fechaDeVencimiento.Split('/');

                if (int.Parse(datos[0]) > int.Parse(datos2[0]))
                {
                  
                    fechaCorrectaDeVencimiento = fechaDeEmision;
                }

            }
            else
            {
                fechaCorrectaDeVencimiento = fechaDeEmision;
            }
            
            return fechaCorrectaDeVencimiento;
        }


        public String convertirAFechaValidaParaNCEEConFolio(String fechaAConvertir)
        {
            String fechaValida = "";
            if (String.IsNullOrEmpty(fechaAConvertir) != true)
            {
                string[] datos = fechaAConvertir.Split('/');

                fechaValida = datos[1] + "/" + datos[2] + "/" + datos[0];
            }
            else
            {
                fechaValida = "Fecha de vencimiento ausente en factura";
            }
            return fechaValida;

        }





        public String calcularFechaDeVencimientoDeCopec(String fechaDeEmision)
        {
            String fechaVencimientoDeFactura = "Esta debiese ser una fecha de vencimientoDeCopec";


            string[] datos = fechaDeEmision.Split('/');

            int mes = int.Parse(datos[1]);
            int anio = int.Parse(datos[0]);
            if (mes == 12)
            {
                anio = anio + 1;
                mes = 01;
            }
            else
            {
                mes = mes + 1;
            }

            fechaVencimientoDeFactura = datos[2] + "/" + mes.ToString() + "/" + anio.ToString();

            return fechaVencimientoDeFactura;

        }

        public String determinarCodigoDeProducto(String rutCliente, String tipoDeFactura, String codigoDelProducto)
        {

            // los ruts de CCU son:
            // 91041000-8
            //96989120-4
            //99501760-1
            //99554560-8
            //99586280-8

            //el rut de COPEC es:
            //99520000-7, y se supone que es el único que tiene que tener
            //tratamiento especial


            //el rut de Entel es:
            //92580000-7 y al 20 de Junio del 2022, solo debiese emitir 4 facturas; una por cada centro de costo donde hay internet fijo.
            //Tenemos internet fijo en Renca en ANTONIO MACEO (CALLE) 2693, Rancagua (CCU)
            //en RUTA H-30 (CALLE) 2333, Talca (no CCU) en LONGITUDINAL SUR KM (CARR) 245  y la central
            //el campo DirRecep del XML indica la direccion

            String codigoDeProducto = "";
            if (rutCliente== "91041000-8" || rutCliente == "96989120-4" || rutCliente == "99501760-1" || rutCliente == "99554560-8" || rutCliente == "99586280-8")
            {
                if (tipoDeFactura == "FCEE")
                {
                    codigoDeProducto = "420724E";
                }
                else
                {
                    codigoDeProducto = "420724";
                }
            }
            else
            {
                codigoDeProducto = codigoDelProducto;
            }

            return codigoDeProducto;
        }

        public Boolean determinarSiEsRutDeCCU(String rutAVerificar)
        {
            Boolean esRutDeCCU = false;


            switch (rutAVerificar)
            {
                case "91041000-8":
                    esRutDeCCU = true;
                    break;
                case "96989120-4":
                    esRutDeCCU = true;
                    break;
                case "99501760-1":
                    esRutDeCCU = true;
                    break;
                case "99554560-8":
                    esRutDeCCU = true;
                    break;
                case "99586280-8":
                    esRutDeCCU = true;
                    break;
                default:
                    esRutDeCCU = false;
                    break;
            }

            return esRutDeCCU;

        }

        public String direccionSiEsQueEnManagerNoTiene(String rut)
        {
            String dir = "Casa Matriz";

            if(rut== "19043336-6" || rut == "79936340-2" || rut == "77178392-9")
            {
                dir = "";
            }

            return dir;
        }   

       


        private Factura validarMontoExentoDeFacturaNoContabilizada(Factura fc)
        {
            Factura fcActualizada = new Factura();

            int totalDocumento = int.Parse(fc.TotalDelDocumento);

            int afecto = int.Parse(fc.MontoAfecto);
            int exento = int.Parse(fc.MontoExento);
            int iva = int.Parse(fc.MontoIva);

            int posibleTotalDelDocumento = afecto + exento + iva;

            //si es una factura afecta o una nota de credito y además el total del documento no cuadra con la suma de afecto, exento e iva
            //entonces a la factura le falta algún valor exento
            if ((fc.TipoDeDocumento == "FACE" || fc.TipoDeDocumento == "NCCE") && (totalDocumento != posibleTotalDelDocumento))
            {
                int exentoFaltante = totalDocumento - posibleTotalDelDocumento;
                exento = exento + exentoFaltante;


                fcActualizada.TipoDeDocumento = fc.TipoDeDocumento;
                fcActualizada.NumeroDelDocumento = fc.NumeroDelDocumento;
                fcActualizada.FechaDeDocumento = fc.FechaDeDocumento;
                fcActualizada.FechaContableDeDocumento = fc.FechaContableDeDocumento;
                fcActualizada.FechaDeVencimientoDeDocumento = fc.FechaDeVencimientoDeDocumento;
                fcActualizada.CodigoDeUnidadDeNegocio = fc.CodigoDeUnidadDeNegocio;
                fcActualizada.RutCliente = fc.RutCliente;
                fcActualizada.DireccionDelCliente = fc.DireccionDelCliente;
                fcActualizada.RutFacturador = fc.RutFacturador;
                fcActualizada.CodigoVendedor = fc.CodigoVendedor;
                fcActualizada.CodigoComisionista = fc.CodigoComisionista;
                fcActualizada.Probabilidad = fc.Probabilidad;
                fcActualizada.ListaPrecio = fc.ListaPrecio;
                fcActualizada.PlazoPago = fc.PlazoPago;
                fcActualizada.MonedaDelDocumento = fc.MonedaDelDocumento;
                fcActualizada.TasaDeCambio = fc.TasaDeCambio;
                fcActualizada.MontoAfecto = fc.MontoAfecto;
                fcActualizada.MontoExento = exento.ToString();
                fcActualizada.MontoIva = fc.MontoIva;
                fcActualizada.MontoImpuestosEspecificos = fc.MontoImpuestosEspecificos;
                fcActualizada.MontoIvaRetenido = fc.MontoIvaRetenido;
                fcActualizada.MontoImpuestosRetenidos = fc.MontoImpuestosRetenidos;
                fcActualizada.TipoDeDescuentoGlobal = fc.TipoDeDescuentoGlobal;
                fcActualizada.DescuentoGlobal = fc.DescuentoGlobal;
                fcActualizada.TotalDelDocumento = fc.TotalDelDocumento;
                fcActualizada.DeudaPendiente = fc.DeudaPendiente;
                fcActualizada.TipoDocReferencia = fc.TipoDocReferencia;
                fcActualizada.NumDocReferencia = fc.NumDocReferencia;
                fcActualizada.FechaDocReferencia = fc.FechaDocReferencia;
                fcActualizada.CodigoDelProducto = fc.CodigoDelProducto;
                fcActualizada.Cantidad = fc.Cantidad;
                fcActualizada.Unidad = fc.Unidad;
                fcActualizada.PrecioUnitario = fc.PrecioUnitario;
                fcActualizada.MonedaDelDetalle = fc.MonedaDelDetalle;
                fcActualizada.TasaDeCambio2 = fc.TasaDeCambio2;
                fcActualizada.NumeroDeSerie = fc.NumeroDeSerie;
                fcActualizada.NumeroDeLote = fc.NumeroDeLote;
                fcActualizada.FechaDeVencimiento = fc.FechaDeVencimiento;
                fcActualizada.CentroDeCostos = fc.CentroDeCostos;
                fcActualizada.TipoDeDescuento = fc.TipoDeDescuento;
                fcActualizada.Descuento = fc.Descuento;
                fcActualizada.Ubicacion = fc.Ubicacion;
                fcActualizada.Bodega = fc.Bodega;
                fcActualizada.Concepto1 = fc.Concepto1;
                fcActualizada.Concepto2 = fc.Concepto2;
                fcActualizada.Concepto3 = fc.Concepto3;
                fcActualizada.Concepto4 = fc.Concepto4;
                fcActualizada.Descripcion = fc.Descripcion;
                fcActualizada.DescripcionAdicional = fc.DescripcionAdicional;
                fcActualizada.Stock = fc.Stock;
                fcActualizada.Comentario11 = fc.Comentario11;
                fcActualizada.Comentario21 = fc.Comentario21;
                fcActualizada.Comentario31 = fc.Comentario31;
                fcActualizada.Comentario41 = fc.Comentario41;
                fcActualizada.Comentario51 = fc.Comentario51;
                fcActualizada.CodigoImpuestoEspecifico1 = fc.CodigoImpuestoEspecifico1;
                fcActualizada.MontoImpuestoEspecifico1 = fc.MontoImpuestoEspecifico1;
                fcActualizada.CodigoImpuestoEspecifico2 = fc.CodigoImpuestoEspecifico2;
                fcActualizada.MontoImpuestoEspecifico2 = fc.MontoImpuestoEspecifico2;
                fcActualizada.Modalidad = fc.Modalidad;
                fcActualizada.Glosa = fc.Glosa;
                fcActualizada.Referencia = fc.Referencia;
                fcActualizada.FechaDeComprometida = fc.FechaDeComprometida;
                fcActualizada.PorcentajeCEEC = "";
                fcActualizada.ImpuestoLey18211 = "";
                fcActualizada.IvaLey18211 = "";
                fcActualizada.CodigoKitFlexible = fc.CodigoKitFlexible;
                fcActualizada.AjusteIva = fc.AjusteIva;

                return fcActualizada;
            }
            else
            {
                return fc;
            }


        }




        private FacturaContabilizada validarMontoExentoDeFacturaContabilizada(FacturaContabilizada fc)
        {
            FacturaContabilizada fcActualizada = new FacturaContabilizada();

            int totalDocumento = int.Parse(fc.TotalDelDocumento);

            int afecto = int.Parse(fc.MontoAfecto);
            int exento = int.Parse(fc.MontoExento);
            int iva = int.Parse(fc.MontoIva);

            int posibleTotalDelDocumento = afecto + exento + iva;

            //si es una factura afecta o una nota de credito y además el total del documento no cuadra con la suma de afecto, exento e iva
            //entonces a la factura le falta algún valor exento
            if ((fc.TipoDeDocumento == "FACE" || fc.TipoDeDocumento == "NCCE") && (totalDocumento != posibleTotalDelDocumento))
            {
                int exentoFaltante = totalDocumento - posibleTotalDelDocumento;
                exento = exento + exentoFaltante;


                fcActualizada.TipoDeDocumento = fc.TipoDeDocumento;
                fcActualizada.NumeroDelDocumento = fc.NumeroDelDocumento;
                fcActualizada.FechaDeDocumento = fc.FechaDeDocumento;
                fcActualizada.FechaContableDeDocumento = fc.FechaContableDeDocumento;
                fcActualizada.FechaDeVencimientoDeDocumento = fc.FechaDeVencimientoDeDocumento;
                fcActualizada.CodigoDeUnidadDeNegocio = fc.CodigoDeUnidadDeNegocio;
                fcActualizada.RutCliente = fc.RutCliente;
                fcActualizada.DireccionDelCliente = fc.DireccionDelCliente;
                fcActualizada.RutFacturador = fc.RutFacturador;
                fcActualizada.CodigoVendedor = fc.CodigoVendedor;
                fcActualizada.CodigoComisionista = fc.CodigoComisionista;
                fcActualizada.Probabilidad = fc.Probabilidad;
                fcActualizada.ListaPrecio = fc.ListaPrecio;
                fcActualizada.PlazoPago = fc.PlazoPago;
                fcActualizada.MonedaDelDocumento = fc.MonedaDelDocumento;
                fcActualizada.TasaDeCambio = fc.TasaDeCambio;
                fcActualizada.MontoAfecto = fc.MontoAfecto;
                fcActualizada.MontoExento = exento.ToString();
                fcActualizada.MontoIva = fc.MontoIva;
                fcActualizada.MontoImpuestosEspecificos = fc.MontoImpuestosEspecificos;
                fcActualizada.MontoIvaRetenido = fc.MontoIvaRetenido;
                fcActualizada.MontoImpuestosRetenidos = fc.MontoImpuestosRetenidos;
                fcActualizada.TipoDeDescuentoGlobal = fc.TipoDeDescuentoGlobal;
                fcActualizada.DescuentoGlobal = fc.DescuentoGlobal;
                fcActualizada.TotalDelDocumento = fc.TotalDelDocumento;
                fcActualizada.DeudaPendiente = fc.DeudaPendiente;
                fcActualizada.TipoDocReferencia = fc.TipoDocReferencia;
                fcActualizada.NumDocReferencia = fc.NumDocReferencia;
                fcActualizada.FechaDocReferencia = fc.FechaDocReferencia;
                fcActualizada.CodigoDelProducto = fc.CodigoDelProducto;
                fcActualizada.Cantidad = fc.Cantidad;
                fcActualizada.Unidad = fc.Unidad;
                fcActualizada.PrecioUnitario = fc.PrecioUnitario;
                fcActualizada.MonedaDelDetalle = fc.MonedaDelDetalle;
                fcActualizada.TasaDeCambio2 = fc.TasaDeCambio2;
                fcActualizada.NumeroDeSerie = fc.NumeroDeSerie;
                fcActualizada.NumeroDeLote = fc.NumeroDeLote;
                fcActualizada.FechaDeVencimiento = fc.FechaDeVencimiento;
                fcActualizada.CentroDeCostos = fc.CentroDeCostos;
                fcActualizada.TipoDeDescuento = fc.TipoDeDescuento;
                fcActualizada.Descuento = fc.Descuento;
                fcActualizada.Ubicacion = fc.Ubicacion;
                fcActualizada.Bodega = fc.Bodega;
                fcActualizada.Concepto1 = fc.Concepto1;
                fcActualizada.Concepto2 = fc.Concepto2;
                fcActualizada.Concepto3 = fc.Concepto3;
                fcActualizada.Concepto4 = fc.Concepto4;
                fcActualizada.Descripcion = fc.Descripcion;
                fcActualizada.DescripcionAdicional = fc.DescripcionAdicional;
                fcActualizada.Stock = fc.Stock;
                fcActualizada.Comentario11 = fc.Comentario11;
                fcActualizada.Comentario21 = fc.Comentario21;
                fcActualizada.Comentario31 = fc.Comentario31;
                fcActualizada.Comentario41 = fc.Comentario41;
                fcActualizada.Comentario51 = fc.Comentario51;
                fcActualizada.CodigoImpuestoEspecifico1 = fc.CodigoImpuestoEspecifico1;
                fcActualizada.MontoImpuestoEspecifico1 = fc.MontoImpuestoEspecifico1;
                fcActualizada.CodigoImpuestoEspecifico2 = fc.CodigoImpuestoEspecifico2;
                fcActualizada.MontoImpuestoEspecifico2 = fc.MontoImpuestoEspecifico2;
                fcActualizada.Modalidad = fc.Modalidad;
                fcActualizada.Glosa = fc.Glosa;
                fcActualizada.Referencia = fc.Referencia;
                fcActualizada.FechaDeComprometida = fc.FechaDeComprometida;
                fcActualizada.PorcentajeCEEC = "";
                fcActualizada.ImpuestoLey18211 = "";
                fcActualizada.IvaLey18211 = "";
                fcActualizada.CodigoKitFlexible = fc.CodigoKitFlexible;
                fcActualizada.AjusteIva = fc.AjusteIva;

                return fcActualizada;
            }
            else
            {
                return fc;
            }


        }

        private void button6_Click(object sender, EventArgs e)
        {
            //aqui habria que tomar un excel que tenga 2 hojas, la primera con las facturas de CCU, y la segunda con el centro que les corresponde


            string sFileName = "";

            OpenFileDialog choofdlog = new OpenFileDialog();
            choofdlog.Filter = "Archivos XLSX (*.xlsx)|*.xlsx";
            choofdlog.FilterIndex = 1;
            choofdlog.Multiselect = true;

            string[] arrAllFiles = new string[] { };

            MessageBox.Show("Seleccionar excel de facturas de manager (debe tener 2 hojas)");

            while (true)
            {
                if (choofdlog.ShowDialog() == DialogResult.OK)
                {
                    sFileName = choofdlog.FileName;
                    arrAllFiles = choofdlog.FileNames; //used when Multiselect = true           
                    break;
                }
                else
                {
                    MessageBox.Show("Debe seleccionar un Excel válido. Cerrando programa");
                    System.Environment.Exit(0);
                }

            }



            List<RegistroCruzadoConInformacionDeCCU> excelCruzadoConInfoDeCCU = leerExcelDeFacturasCCUACruzar(sFileName);

            String pathDeDescargas = getCarpetaDeDescargas() + "" + @"\Facturas de CCU costeadas.xlsx";
            var archivo = new FileInfo(pathDeDescargas);
            SaveExcelFileCosteoFacturasCCU(excelCruzadoConInfoDeCCU, archivo);
            MessageBox.Show("Se creo archivo de facturas de CCU costeadas en carpeta de descargas");
        }





        private List<RegistroCruzadoConInformacionDeCCU> leerExcelDeFacturasCCUACruzar(String filePath)
        {
            List<RegistroCruzadoConInformacionDeCCU> listadoDeFacturasCruzadasConInfoDeCCU = new List<RegistroCruzadoConInformacionDeCCU>();
            List<RegistroDeCCU> listadoDeRegistrosDeCCU = new List<RegistroDeCCU>();

            FileInfo existingFile = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int colCount = worksheet.Dimension.End.Column;  //get Column Count
                int rowCount = worksheet.Dimension.End.Row;     //get row count
          
                for (int row = 1; row <= rowCount; row++)
                {

                    RegistroCruzadoConInformacionDeCCU rccic = new RegistroCruzadoConInformacionDeCCU();
                    rccic.TipoDeDocumento = worksheet.Cells[row, 1].Value?.ToString().Trim();
                    rccic.NumeroDelDocumento = worksheet.Cells[row, 2].Value?.ToString().Trim();
                    rccic.FechaDeDocumento = worksheet.Cells[row, 3].Value?.ToString().Trim();
                    rccic.FechaContableDeDocumento = worksheet.Cells[row, 4].Value?.ToString().Trim();
                    rccic.FechaDeVencimientoDeDocumento = worksheet.Cells[row, 5].Value?.ToString().Trim();
                    rccic.CodigoDeUnidadDeNegocio = worksheet.Cells[row, 6].Value?.ToString().Trim();
                    rccic.RutCliente = worksheet.Cells[row, 7].Value?.ToString().Trim();
                    rccic.DireccionDelCliente = worksheet.Cells[row, 8].Value?.ToString().Trim();
                    rccic.RutFacturador = worksheet.Cells[row, 9].Value?.ToString().Trim();
                    rccic.CodigoVendedor = worksheet.Cells[row, 10].Value?.ToString().Trim();
                    rccic.CodigoComisionista = worksheet.Cells[row, 11].Value?.ToString().Trim();
                    rccic.Probabilidad = worksheet.Cells[row, 12].Value?.ToString().Trim();
                    rccic.ListaPrecio = worksheet.Cells[row, 13].Value?.ToString().Trim();
                    rccic.PlazoPago = worksheet.Cells[row, 14].Value?.ToString().Trim();
                    rccic.MonedaDelDocumento = worksheet.Cells[row, 15].Value?.ToString().Trim();
                    rccic.TasaDeCambio = worksheet.Cells[row, 16].Value?.ToString().Trim();
                    rccic.MontoAfecto = worksheet.Cells[row, 17].Value?.ToString().Trim();
                    rccic.MontoExento = worksheet.Cells[row, 18].Value?.ToString().Trim();
                    rccic.MontoIva = worksheet.Cells[row, 19].Value?.ToString().Trim();
                    rccic.MontoImpuestosEspecificos = worksheet.Cells[row, 20].Value?.ToString().Trim();
                    rccic.MontoIvaRetenido = worksheet.Cells[row, 21].Value?.ToString().Trim();
                    rccic.MontoImpuestosRetenidos = worksheet.Cells[row, 22].Value?.ToString().Trim();
                    rccic.TipoDeDescuentoGlobal = worksheet.Cells[row, 23].Value?.ToString().Trim();
                    rccic.DescuentoGlobal = worksheet.Cells[row, 24].Value?.ToString().Trim();
                    rccic.TotalDelDocumento = worksheet.Cells[row, 25].Value?.ToString().Trim();
                    rccic.DeudaPendiente = worksheet.Cells[row, 26].Value?.ToString().Trim();
                    rccic.TipoDocReferencia = worksheet.Cells[row, 27].Value?.ToString().Trim();
                    rccic.NumDocReferencia = worksheet.Cells[row, 28].Value?.ToString().Trim();
                    rccic.FechaDocReferencia = worksheet.Cells[row, 29].Value?.ToString().Trim();
                    rccic.CodigoDelProducto = worksheet.Cells[row, 30].Value?.ToString().Trim();
                    rccic.Cantidad = worksheet.Cells[row, 31].Value?.ToString().Trim();
                    rccic.Unidad = worksheet.Cells[row, 32].Value?.ToString().Trim();
                    rccic.PrecioUnitario = worksheet.Cells[row, 33].Value?.ToString().Trim();
                    rccic.MonedaDelDetalle = worksheet.Cells[row, 34].Value?.ToString().Trim();
                    rccic.TasaDeCambio2 = worksheet.Cells[row, 35].Value?.ToString().Trim();
                    rccic.NumeroDeSerie = worksheet.Cells[row, 36].Value?.ToString().Trim();
                    rccic.NumeroDeLote = worksheet.Cells[row, 37].Value?.ToString().Trim();
                    rccic.FechaDeVencimiento = worksheet.Cells[row, 38].Value?.ToString().Trim();
                    rccic.CentroDeCostos = worksheet.Cells[row, 39].Value?.ToString().Trim();
                    rccic.TipoDeDescuento = worksheet.Cells[row, 40].Value?.ToString().Trim();
                    rccic.Descuento = worksheet.Cells[row, 41].Value?.ToString().Trim();
                    rccic.Ubicacion = worksheet.Cells[row, 42].Value?.ToString().Trim();
                    rccic.Bodega = worksheet.Cells[row, 43].Value?.ToString().Trim();
                    rccic.Concepto1 = worksheet.Cells[row, 44].Value?.ToString().Trim();
                    rccic.Concepto2 = worksheet.Cells[row, 45].Value?.ToString().Trim();
                    rccic.Concepto3 = worksheet.Cells[row, 46].Value?.ToString().Trim();
                    rccic.Concepto4 = worksheet.Cells[row, 47].Value?.ToString().Trim();
                    rccic.Descripcion = worksheet.Cells[row, 48].Value?.ToString().Trim();
                    rccic.DescripcionAdicional = worksheet.Cells[row, 49].Value?.ToString().Trim();
                    rccic.Stock = worksheet.Cells[row, 50].Value?.ToString().Trim();
                    rccic.Comentario11 = worksheet.Cells[row, 51].Value?.ToString().Trim();
                    rccic.Comentario21 = worksheet.Cells[row, 52].Value?.ToString().Trim();
                    rccic.Comentario31 = worksheet.Cells[row, 53].Value?.ToString().Trim();
                    rccic.Comentario41 = worksheet.Cells[row, 54].Value?.ToString().Trim();
                    rccic.Comentario51 = worksheet.Cells[row, 55].Value?.ToString().Trim();
                    rccic.CodigoImpuestoEspecifico1 = worksheet.Cells[row, 56].Value?.ToString().Trim();
                    rccic.MontoImpuestoEspecifico1 = worksheet.Cells[row, 57].Value?.ToString().Trim();
                    rccic.CodigoImpuestoEspecifico2 = worksheet.Cells[row, 58].Value?.ToString().Trim();
                    rccic.MontoImpuestoEspecifico2 = worksheet.Cells[row, 59].Value?.ToString().Trim();
                    rccic.Modalidad = worksheet.Cells[row, 60].Value?.ToString().Trim();
                    rccic.Glosa = worksheet.Cells[row, 61].Value?.ToString().Trim();
                    rccic.Referencia = worksheet.Cells[row, 62].Value?.ToString().Trim();
                    rccic.FechaDeComprometida = worksheet.Cells[row, 63].Value?.ToString().Trim();
                    rccic.PorcentajeCEEC = worksheet.Cells[row, 64].Value?.ToString().Trim();
                    rccic.ImpuestoLey18211 = worksheet.Cells[row, 65].Value?.ToString().Trim();
                    rccic.IvaLey18211 = worksheet.Cells[row, 66].Value?.ToString().Trim();
                    rccic.CodigoKitFlexible = worksheet.Cells[row, 67].Value?.ToString().Trim();
                    rccic.AjusteIva = worksheet.Cells[row, 68].Value?.ToString().Trim();

                    listadoDeFacturasCruzadasConInfoDeCCU.Add(rccic);

                }

                ExcelWorksheet worksheet2 = package.Workbook.Worksheets[1];
                int colCount2 = worksheet2.Dimension.End.Column;  //get Column Count
                int rowCount2 = worksheet2.Dimension.End.Row;     //get row count


                for (int row = 1; row <= rowCount2; row++)
                {

                    RegistroDeCCU rCCU = new RegistroDeCCU();
                    rCCU.Rut = worksheet2.Cells[row, 1].Value?.ToString().Trim();
                    rCCU.Folio = worksheet2.Cells[row, 2].Value?.ToString().Trim();
                    rCCU.Centro = worksheet2.Cells[row, 3].Value?.ToString().Trim();

                    listadoDeRegistrosDeCCU.Add(rCCU);

                }



                foreach (var item in listadoDeRegistrosDeCCU)
                {
                    foreach (var item2 in listadoDeFacturasCruzadasConInfoDeCCU)
                    {

                        if ( (item.Rut==item2.RutCliente) && (item.Folio == item2.NumeroDelDocumento))
                        {

                            item2.Glosa = "COSTEADO CON INFO DE CCU";

                            //203 / 303   Administracion
                            //204 / 304   Interplantas
                            //208 / 308   Emprendedores
                            //205 / 305   Illapel
                            //207 / 307   San Antonio
                            //200 / 300   Melipilla
                            //206 / 306   Santiago
                            //201 / 301   Rancagua
                            //202 / 302   Curico

                            //switch para determinar a que centro de costo va
                            switch (item.Centro)
                            {
                                case "ADMINISTRACION":
                                    item2.CentroDeCostos = "203";
                                    break;
                                case "INTERPLANTA":
                                    item2.CentroDeCostos = "204";
                                    item2.CodigoDeUnidadDeNegocio = "2";
                                    break;
                                case "EMPRENDEDORES":
                                    item2.CentroDeCostos = "208";
                                    break;
                                case "ILLAPEL":
                                    item2.CentroDeCostos = "205";
                                    break;
                                case "SAN ANTONIO":
                                    item2.CentroDeCostos = "207";
                                    break;
                                case "MELIPILLA":
                                    item2.CentroDeCostos = "200";
                                    break;
                                case "SANTIAGO":
                                    item2.CentroDeCostos = "206";
                                    break;
                                case "RANCAGUA":
                                    item2.CentroDeCostos = "201";
                                    break;
                                case "CURICO":
                                    item2.CentroDeCostos = "202";
                                    break;
                                case "#N/D":
                                    item2.CentroDeCostos = "209";
                                    break;
                                default:
                                    item2.CentroDeCostos = item2.CentroDeCostos;
                                    break;

                            }


                        }


                    }

                }

            }
            return listadoDeFacturasCruzadasConInfoDeCCU;

        }

        private void button7_Click(object sender, EventArgs e)
        {
            Console.WriteLine("Se llama a funcion para costear detalles de facturas NO CCU");


            //hay que tomar un Excel con 2 hojas; la primera con las facturas a costear, la segunda con el costeo de estas facturas

            string sFileName = "";

            OpenFileDialog choofdlog = new OpenFileDialog();
            choofdlog.Filter = "Archivos XLSX (*.xlsx)|*.xlsx";
            choofdlog.FilterIndex = 1;
            choofdlog.Multiselect = true;

            string[] arrAllFiles = new string[] { };

            MessageBox.Show("Seleccionar excel de facturas NO CCU (debe tener 2 hojas)");
            while (true)
            {


                if (choofdlog.ShowDialog() == DialogResult.OK)
                {
                    sFileName = choofdlog.FileName;
                    arrAllFiles = choofdlog.FileNames; //used when Multiselect = true
                    break;
                }
                else
                {
                    MessageBox.Show("Debe seleccionar un Excel válido. Cerrando programa");
                    System.Environment.Exit(0);
                }


            }


            List<Factura> listadoDeFacturasCosteadas = leerExcelDeFacturasNOCCUACostear(sFileName);

            int cantidadDeFacturasACostear = contarFacturasPresentesEnCosteo(sFileName);

            String pathDeDescargas = getCarpetaDeDescargas() + "" + @"\Facturas NO CCU costeadas ("+ cantidadDeFacturasACostear + " facturas costeadas).xlsx";
            var archivo = new FileInfo(pathDeDescargas);
            
            
            SaveExcelFileCosteoFacturasNOCCU(listadoDeFacturasCosteadas, archivo);
            MessageBox.Show("Se creo archivo de facturas NO CCU costeadas");



        }

        private static async Task SaveExcelFileCosteoFacturasNOCCU(List<Factura> facturasNOCCUCosteadas, FileInfo file)
        {
            var package = new ExcelPackage(file);
            var ws = package.Workbook.Worksheets.Add("Facturas NO CCU costeadas");

            var range = ws.Cells["A1"].LoadFromCollection(facturasNOCCUCosteadas, true);

            range.AutoFitColumns();

            await package.SaveAsync();
        }


        private int contarFacturasPresentesEnCosteo(String filePath)
        {
            int cantidadDeFacturas = 0;

            List<String> listadoDeFacturasACostear = new List<String>();

            FileInfo existingFile = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {

                ExcelWorksheet hojaDeCosteos = package.Workbook.Worksheets[1];
                int colCountCosteos = hojaDeCosteos.Dimension.End.Column; 
                int rowCountCosteos = hojaDeCosteos.Dimension.End.Row;

                for (int row = 1; row <= rowCountCosteos; row++)
                {

                    CosteoDeFacturaNOCCU costeoDeFactura = new CosteoDeFacturaNOCCU();
                    costeoDeFactura.Folio = hojaDeCosteos.Cells[row, 3].Value?.ToString().Trim();
                    costeoDeFactura.Rut = hojaDeCosteos.Cells[row, 1].Value?.ToString().Trim();
            

                    if (costeoDeFactura.Rut != "rut")
                    {
                        listadoDeFacturasACostear.Add(costeoDeFactura.Rut + costeoDeFactura.Folio);
                    }


                }

            }


            listadoDeFacturasACostear = listadoDeFacturasACostear.Distinct().ToList();

            cantidadDeFacturas = listadoDeFacturasACostear.Count;


            return cantidadDeFacturas;

        }

        private List<Factura> leerExcelDeFacturasNOCCUACostear(String filePath)
        {
            List<Factura> facturasLeidasEnPrimeraHoja = new List<Factura>();
            List<CosteoDeFacturaNOCCU> listadoDeCosteos = new List<CosteoDeFacturaNOCCU>();

           

            List<Factura> listadoDeFacturasCosteadas = new List<Factura>();

            FileInfo existingFile = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                //facturas
                ExcelWorksheet hojaDeFacturas = package.Workbook.Worksheets[0];
                int colCountFacturas = hojaDeFacturas.Dimension.End.Column;  //get Column Count
                int rowCountFacturas = hojaDeFacturas.Dimension.End.Row;     //get row count


                for (int row = 1; row <= rowCountFacturas; row++)
                {

                    Factura rccic = new Factura();

                    rccic.TipoDeDocumento = hojaDeFacturas.Cells[row, 1].Value?.ToString().Trim();
                    rccic.NumeroDelDocumento = hojaDeFacturas.Cells[row, 2].Value?.ToString().Trim();
                    rccic.FechaDeDocumento = hojaDeFacturas.Cells[row, 3].Value?.ToString().Trim();
                    rccic.FechaContableDeDocumento = hojaDeFacturas.Cells[row, 4].Value?.ToString().Trim();
                    rccic.FechaDeVencimientoDeDocumento = hojaDeFacturas.Cells[row, 5].Value?.ToString().Trim();
                    rccic.CodigoDeUnidadDeNegocio = hojaDeFacturas.Cells[row, 6].Value?.ToString().Trim();
                    rccic.RutCliente = hojaDeFacturas.Cells[row, 7].Value?.ToString().Trim();
                    rccic.DireccionDelCliente = hojaDeFacturas.Cells[row, 8].Value?.ToString().Trim();
                    rccic.RutFacturador = hojaDeFacturas.Cells[row, 9].Value?.ToString().Trim();
                    rccic.CodigoVendedor = hojaDeFacturas.Cells[row, 10].Value?.ToString().Trim();
                    rccic.CodigoComisionista = hojaDeFacturas.Cells[row, 11].Value?.ToString().Trim();
                    rccic.Probabilidad = hojaDeFacturas.Cells[row, 12].Value?.ToString().Trim();
                    rccic.ListaPrecio = hojaDeFacturas.Cells[row, 13].Value?.ToString().Trim();
                    rccic.PlazoPago = hojaDeFacturas.Cells[row, 14].Value?.ToString().Trim();
                    rccic.MonedaDelDocumento = hojaDeFacturas.Cells[row, 15].Value?.ToString().Trim();
                    rccic.TasaDeCambio = hojaDeFacturas.Cells[row, 16].Value?.ToString().Trim();
                    rccic.MontoAfecto = hojaDeFacturas.Cells[row, 17].Value?.ToString().Trim();
                    rccic.MontoExento = hojaDeFacturas.Cells[row, 18].Value?.ToString().Trim();
                    rccic.MontoIva = hojaDeFacturas.Cells[row, 19].Value?.ToString().Trim();
                    rccic.MontoImpuestosEspecificos = hojaDeFacturas.Cells[row, 20].Value?.ToString().Trim();
                    rccic.MontoIvaRetenido = hojaDeFacturas.Cells[row, 21].Value?.ToString().Trim();
                    rccic.MontoImpuestosRetenidos = hojaDeFacturas.Cells[row, 22].Value?.ToString().Trim();
                    rccic.TipoDeDescuentoGlobal = hojaDeFacturas.Cells[row, 23].Value?.ToString().Trim();
                    rccic.DescuentoGlobal = hojaDeFacturas.Cells[row, 24].Value?.ToString().Trim();
                    rccic.TotalDelDocumento = hojaDeFacturas.Cells[row, 25].Value?.ToString().Trim();
                    rccic.DeudaPendiente = hojaDeFacturas.Cells[row, 26].Value?.ToString().Trim();
                    rccic.TipoDocReferencia = hojaDeFacturas.Cells[row, 27].Value?.ToString().Trim();
                    rccic.NumDocReferencia = hojaDeFacturas.Cells[row, 28].Value?.ToString().Trim();
                    rccic.FechaDocReferencia = hojaDeFacturas.Cells[row, 29].Value?.ToString().Trim();
                    rccic.CodigoDelProducto = hojaDeFacturas.Cells[row, 30].Value?.ToString().Trim();
                    rccic.Cantidad = hojaDeFacturas.Cells[row, 31].Value?.ToString().Trim();
                    rccic.Unidad = hojaDeFacturas.Cells[row, 32].Value?.ToString().Trim();
                    rccic.PrecioUnitario = hojaDeFacturas.Cells[row, 33].Value?.ToString().Trim();
                    rccic.MonedaDelDetalle = hojaDeFacturas.Cells[row, 34].Value?.ToString().Trim();
                    rccic.TasaDeCambio2 = hojaDeFacturas.Cells[row, 35].Value?.ToString().Trim();
                    rccic.NumeroDeSerie = hojaDeFacturas.Cells[row, 36].Value?.ToString().Trim();
                    rccic.NumeroDeLote = hojaDeFacturas.Cells[row, 37].Value?.ToString().Trim();
                    rccic.FechaDeVencimiento = hojaDeFacturas.Cells[row, 38].Value?.ToString().Trim();
                    rccic.CentroDeCostos = hojaDeFacturas.Cells[row, 39].Value?.ToString().Trim();
                    rccic.TipoDeDescuento = hojaDeFacturas.Cells[row, 40].Value?.ToString().Trim();
                    rccic.Descuento = hojaDeFacturas.Cells[row, 41].Value?.ToString().Trim();
                    rccic.Ubicacion = hojaDeFacturas.Cells[row, 42].Value?.ToString().Trim();
                    rccic.Bodega = hojaDeFacturas.Cells[row, 43].Value?.ToString().Trim();
                    rccic.Concepto1 = hojaDeFacturas.Cells[row, 44].Value?.ToString().Trim();
                    rccic.Concepto2 = hojaDeFacturas.Cells[row, 45].Value?.ToString().Trim();
                    rccic.Concepto3 = hojaDeFacturas.Cells[row, 46].Value?.ToString().Trim();
                    rccic.Concepto4 = hojaDeFacturas.Cells[row, 47].Value?.ToString().Trim();
                    rccic.Descripcion = hojaDeFacturas.Cells[row, 48].Value?.ToString().Trim();
                    rccic.DescripcionAdicional = hojaDeFacturas.Cells[row, 49].Value?.ToString().Trim();
                    rccic.Stock = hojaDeFacturas.Cells[row, 50].Value?.ToString().Trim();
                    rccic.Comentario11 = hojaDeFacturas.Cells[row, 51].Value?.ToString().Trim();
                    rccic.Comentario21 = hojaDeFacturas.Cells[row, 52].Value?.ToString().Trim();
                    rccic.Comentario31 = hojaDeFacturas.Cells[row, 53].Value?.ToString().Trim();
                    rccic.Comentario41 = hojaDeFacturas.Cells[row, 54].Value?.ToString().Trim();
                    rccic.Comentario51 = hojaDeFacturas.Cells[row, 55].Value?.ToString().Trim();
                    rccic.CodigoImpuestoEspecifico1 = hojaDeFacturas.Cells[row, 56].Value?.ToString().Trim();
                    rccic.MontoImpuestoEspecifico1 = hojaDeFacturas.Cells[row, 57].Value?.ToString().Trim();
                    rccic.CodigoImpuestoEspecifico2 = hojaDeFacturas.Cells[row, 58].Value?.ToString().Trim();
                    rccic.MontoImpuestoEspecifico2 = hojaDeFacturas.Cells[row, 59].Value?.ToString().Trim();
                    rccic.Modalidad = hojaDeFacturas.Cells[row, 60].Value?.ToString().Trim();
                    rccic.Glosa = hojaDeFacturas.Cells[row, 61].Value?.ToString().Trim();
                    rccic.Referencia = hojaDeFacturas.Cells[row, 62].Value?.ToString().Trim();
                    rccic.FechaDeComprometida = hojaDeFacturas.Cells[row, 63].Value?.ToString().Trim();
                    rccic.PorcentajeCEEC = hojaDeFacturas.Cells[row, 64].Value?.ToString().Trim();
                    rccic.ImpuestoLey18211 = hojaDeFacturas.Cells[row, 65].Value?.ToString().Trim();
                    rccic.IvaLey18211 = hojaDeFacturas.Cells[row, 66].Value?.ToString().Trim();
                    rccic.CodigoKitFlexible = hojaDeFacturas.Cells[row, 67].Value?.ToString().Trim();
                    rccic.AjusteIva = hojaDeFacturas.Cells[row, 68].Value?.ToString().Trim();
           
                    facturasLeidasEnPrimeraHoja.Add(rccic);
                    
                                                      

                }




                ExcelWorksheet hojaDeCosteos = package.Workbook.Worksheets[1];
                int colCountCosteos = hojaDeCosteos.Dimension.End.Column;  //get Column Count
                int rowCountCosteos = hojaDeCosteos.Dimension.End.Row;     //get row count

                for (int row = 1; row <= rowCountCosteos; row++)
                {

                    CosteoDeFacturaNOCCU costeoDeFactura = new CosteoDeFacturaNOCCU();
                    costeoDeFactura.Folio = hojaDeCosteos.Cells[row, 3].Value?.ToString().Trim();
                    costeoDeFactura.Rut = hojaDeCosteos.Cells[row, 1].Value?.ToString().Trim();
                    costeoDeFactura.Afecto = hojaDeCosteos.Cells[row, 4].Value?.ToString().Trim();
                    costeoDeFactura.Exento= hojaDeCosteos.Cells[row,5].Value?.ToString().Trim();    
                    costeoDeFactura.CentroDeCosto = hojaDeCosteos.Cells[row, 12].Value?.ToString().Trim();

                    costeoDeFactura.MontoIva = hojaDeCosteos.Cells[row, 6].Value?.ToString().Trim();
                    costeoDeFactura.AjusteIva = hojaDeCosteos.Cells[row, 7].Value?.ToString().Trim();
                    costeoDeFactura.CodigoDelProducto = hojaDeCosteos.Cells[row, 9].Value?.ToString().Trim();
                    costeoDeFactura.Glosa = hojaDeCosteos.Cells[row, 15].Value?.ToString().Trim();// para las observaciones

                    costeoDeFactura.FechaDeDocumento = convertirFechaDePamelaAFechaParaExcelDeManager(hojaDeCosteos.Cells[row, 14].Value?.ToString().Trim());


                    listadoDeCosteos.Add(costeoDeFactura);
           

                }

            }






            List<IdentificadorDeFactura> identificadores = new List<IdentificadorDeFactura>();

            //con las facturas no CCU hay que hacer 2 cosas a la hora de hacer el Excel que las costea
            //1.- Crear tantas filas identicas como costeos haya
            //2.- Alterar valores de precio unitario (debe ser el monto afecto que aparece en el costeo)
            //y centro de costos por cada costeo (debiese venir como letras, el excel se tiene que subir como codigo)
            //3.- Subir a documentos con detalle (contabilizado)

            Boolean existeRegistro = false;

            foreach (var item in facturasLeidasEnPrimeraHoja)
            {
                foreach (var identi in identificadores)
                {
                    if (item.NumeroDelDocumento == identi.Folio && item.RutCliente == identi.Rut)
                    {
                        existeRegistro = true;
                    }

                }


                if (item.NumeroDelDocumento != "NumeroDelDocumento" && item.RutCliente != "RutCliente" && existeRegistro == false)
                {

                    IdentificadorDeFactura i = new IdentificadorDeFactura(item.NumeroDelDocumento, item.RutCliente);
                    identificadores.Add(i);

                }


                existeRegistro = false;

            }


            //Actualizacion 21/07/2022, habría que modificar el programa para que automáticamente genere un registro de factura
            //por cada factura que si este presente en el costeo, pero no en el listado de todas las facturas.

            foreach (var identificador in identificadores)
            {

                foreach (var costeo in listadoDeCosteos)
                {

                    //factura esta presente en ambas hojas
                    if(identificador.Folio==costeo.Folio && identificador.Rut == costeo.Rut && costeo.CodigoDelProducto!="410104")
                    {
                        Factura fc = new Factura();


                        String tipoDeDocumento = "";
                        String fechaDeDocumento = "";
                        String fechaContableDelDocumento = "";
                        String fechaDeVencimientoDelDocumento = "";
                        String codigoDeUnidadDeNegocio = "";

                        String direccionDelCliente = "";

                     
                        String codigoDelProducto = "0";
                        String precioUnitario = "0";

                        String montoAfecto = "0";
                        String montoIva = "0";
                        String montoExento = "0";    
                        String montoTotal = "0";
                        String ajusteIva = "0";
                        String glosa = "";


                        foreach (var facturaLeida in facturasLeidasEnPrimeraHoja)
                        {
                            
                            if(identificador.Folio == facturaLeida.NumeroDelDocumento && identificador.Rut == facturaLeida.RutCliente)
                            {
                                tipoDeDocumento = facturaLeida.TipoDeDocumento;
                                fechaDeDocumento = facturaLeida.FechaDeDocumento;
                                fechaContableDelDocumento = facturaLeida.FechaContableDeDocumento;
                                fechaDeVencimientoDelDocumento = facturaLeida.FechaDeVencimientoDeDocumento;
                                codigoDeUnidadDeNegocio = facturaLeida.CodigoDeUnidadDeNegocio;
                                direccionDelCliente = facturaLeida.DireccionDelCliente;
                                codigoDelProducto = costeo.CodigoDelProducto;

                                //en lo que respecta a precios, solo el precio unitario cambia.
                                precioUnitario = costeo.Afecto;

                                
                                montoAfecto = facturaLeida.MontoAfecto;
                                montoIva = facturaLeida.MontoIva;


                                montoExento = facturaLeida.MontoExento;
                                montoTotal = facturaLeida.TotalDelDocumento;
                                ajusteIva = costeo.AjusteIva;
                                facturaLeida.AjusteIva = costeo.AjusteIva;

                                facturaLeida.Glosa= costeo.Glosa;

                                if (ajusteIva=="")
                                {
                                    ajusteIva = facturaLeida.AjusteIva;
                                }

                                fechaDeDocumento = costeo.FechaDeDocumento;
                                fechaContableDelDocumento = costeo.FechaDeDocumento;
                                fechaDeVencimientoDelDocumento = costeo.FechaDeDocumento;


                            }
                         
                        }

                        // Hasta la columna FechaDocReferencia (AC) todo es totales, así que por cada costeo, 
                        //tiene que haber una fila nueva (ej: 6 costeos, 6 filas).
                        //La columna de código de producto varia según el código entregado por Pamela.
                        //El precio unitario debiese ser el valor afecto de cada costeo
                        //El centro de costo depende de la palabra que viene en el archivo de costeo
                        //El ajuste de iva también depende del valor que viene en la factura
                        //Lo anterior aplica a facturas afectas (FACE), sin impuestos adicionales.
                        //El ajuste de IVA tiene que ser igual en todas las filas de la factura

                        //Si llega a haber un valor Exento, eso tiene el mismo tratamiento que las partes afectas
                        //(el exento va en el precio unitario y se costea al centro apropiado). Importante que el codigo
                        //de producto sea la variacion de exento del producto sujeto al impuesto a ingresar

                        if((fc.TipoDeDocumento=="FACE" || fc.TipoDeDocumento == "NCCE") && fc.MontoExento!="" && fc.MontoExento!="0")
                        {
                            //es una factura con impuestos especiales

                        }else if (fc.TipoDeDocumento=="FCEE")
                        {
                            //es una factura exenta
                        }

                        // a la hora de subir una factura de Copec, se supone que el monto negativo debiese ir en el precio unitario
                        //el afecto es el que dice la factura, iva y total tambien lo que dice en la factura, pero la
                        //suma de los impuestos variables y fijos debiese ir en precio unitario



                        fc.TipoDeDocumento = tipoDeDocumento;
                        fc.NumeroDelDocumento = costeo.Folio;
                        fc.FechaDeDocumento = fechaDeDocumento;
                        fc.FechaContableDeDocumento = fechaContableDelDocumento;
                        fc.FechaDeVencimientoDeDocumento = fechaDeVencimientoDelDocumento;
                        fc.CodigoDeUnidadDeNegocio = codigoDeUnidadDeNegocio;
                        fc.RutCliente = costeo.Rut;
                        fc.DireccionDelCliente = direccionDelCliente;
                        fc.RutFacturador = "";
                        fc.CodigoVendedor = "";
                        fc.CodigoComisionista = "";
                        fc.Probabilidad = "";
                        fc.ListaPrecio = "";
                        fc.PlazoPago = "P01";
                        fc.MonedaDelDocumento = "CLP";
                        fc.TasaDeCambio = "";
                        fc.MontoAfecto = montoAfecto;
                        fc.MontoExento = montoExento;
                        fc.MontoIva = montoIva;
                        fc.MontoImpuestosEspecificos = "";
                        fc.MontoIvaRetenido = "";
                        fc.MontoImpuestosRetenidos = "";
                        fc.TipoDeDescuentoGlobal = "";
                        fc.DescuentoGlobal = "";
                        fc.TotalDelDocumento = montoTotal;
                        fc.DeudaPendiente = fc.TotalDelDocumento;
                        fc.TipoDocReferencia = "";
                        fc.NumDocReferencia = "";
                        fc.FechaDocReferencia = "";
                        fc.CodigoDelProducto = codigoDelProducto;
                        fc.Cantidad = "1";
                        fc.Unidad = "S.U.M";
                        fc.PrecioUnitario = precioUnitario;
                        fc.MonedaDelDetalle = "CLP";
                        fc.TasaDeCambio2 = "1";
                        fc.NumeroDeSerie = "";
                        fc.NumeroDeLote = "";
                        fc.FechaDeVencimiento = "";
                        fc.CentroDeCostos = costeo.CentroDeCosto;


                        switch (fc.CentroDeCostos)
                        {
                            case "ADMINISTRACION":
                                fc.CentroDeCostos = "203";
                                break;
                            case "INTERPLANTA":
                                fc.CentroDeCostos = "204";
                                fc.CodigoDeUnidadDeNegocio = "2";
                                break;
                            case "EMPRENDEDORES":
                                fc.CentroDeCostos = "208";
                                break;
                            case "ILLAPEL":
                                fc.CentroDeCostos = "205";
                                break;
                            case "SAN ANTONIO":
                                fc.CentroDeCostos = "207";
                                break;
                            case "MELIPILLA":
                                fc.CentroDeCostos = "200";
                                break;
                            case "SANTIAGO":
                                fc.CentroDeCostos = "206";
                                break;
                            case "RANCAGUA":
                                fc.CentroDeCostos = "201";
                                break;
                            case "CURICO":
                                fc.CentroDeCostos = "202";
                                break;
                            case "203":
                                fc.CentroDeCostos = "203";
                                break;
                            case "204":
                                fc.CentroDeCostos = "204";
                                fc.CodigoDeUnidadDeNegocio = "2";
                                break;
                            case "208":
                                fc.CentroDeCostos = "208";
                                break;
                            case "205":
                                fc.CentroDeCostos = "205";
                                break;
                            case "207":
                                fc.CentroDeCostos = "207";
                                break;
                            case "200":
                                fc.CentroDeCostos = "200";
                                break;
                            case "206":
                                fc.CentroDeCostos = "206";
                                break;
                            case "201":
                                fc.CentroDeCostos = "201";
                                break;
                            case "202":
                                fc.CentroDeCostos = "202";
                                break;
                            default:
                                fc.CentroDeCostos = "209";
                                break;

                        }



                        fc.TipoDeDescuento = "";
                        fc.Descuento = "";
                        fc.Ubicacion = "";
                        fc.Bodega = "";
                        fc.Concepto1 = "";
                        fc.Concepto2 = "";
                        fc.Concepto3 = "";
                        fc.Concepto4 = "";
                        fc.Descripcion = "";
                        fc.DescripcionAdicional = "";
                        fc.Stock = "0";
                        fc.Comentario11 = "";
                        fc.Comentario21 = "";
                        fc.Comentario31 = "";
                        fc.Comentario41 = "";
                        fc.Comentario51 = "";
                        fc.CodigoImpuestoEspecifico1 = "";
                        fc.MontoImpuestoEspecifico1 = "";
                        fc.CodigoImpuestoEspecifico2 = "";
                        fc.MontoImpuestoEspecifico2 = "";
                        fc.Modalidad = "";
                        fc.Glosa = costeo.Glosa;
                        fc.Referencia = "";
                        fc.FechaDeComprometida = "";
                        fc.PorcentajeCEEC = "";
                        fc.ImpuestoLey18211 = "";
                        fc.IvaLey18211 = "";
                        fc.CodigoKitFlexible = "";
                        fc.AjusteIva = ajusteIva;

                        if (!String.IsNullOrEmpty(fc.RutCliente) && fc.RutCliente!="rut")
                        {
                            listadoDeFacturasCosteadas.Add(fc);
                        }
                    }
                    else if (identificador.Folio == costeo.Folio && identificador.Rut == costeo.Rut && costeo.CodigoDelProducto == "410104")
                    {
                        //esto significa que la factura es de petróleo y que debería aparecer 2 veces, si es que el exento es negativo
                        Factura facturaSinG = new Factura();
                        Factura facturaConG = new Factura();


                        String tipoDeDocumento = "";
                        String fechaDeDocumento = "";
                        String fechaContableDelDocumento = "";
                        String fechaDeVencimientoDelDocumento = "";
                        String codigoDeUnidadDeNegocio = "";

                        String direccionDelCliente = "";


                        String codigoDelProducto = "0";
                        String precioUnitario = "0";

                        String montoAfecto = "0";
                        String montoIva = "0";
                        String montoExento = "0";
                        String montoTotal = "0";
                        String ajusteIva = "0";
                        String glosa = "";


                        foreach (var facturaLeida in facturasLeidasEnPrimeraHoja)
                        {

                            if (identificador.Folio == facturaLeida.NumeroDelDocumento && identificador.Rut == facturaLeida.RutCliente)
                            {
                                tipoDeDocumento = facturaLeida.TipoDeDocumento;
                                fechaDeDocumento = facturaLeida.FechaDeDocumento;
                                fechaContableDelDocumento = facturaLeida.FechaContableDeDocumento;
                                fechaDeVencimientoDelDocumento = facturaLeida.FechaDeVencimientoDeDocumento;
                                codigoDeUnidadDeNegocio = facturaLeida.CodigoDeUnidadDeNegocio;
                                direccionDelCliente = facturaLeida.DireccionDelCliente;
                                codigoDelProducto = costeo.CodigoDelProducto;

                                //en lo que respecta a precios, solo el precio unitario cambia.
                                precioUnitario = costeo.Afecto;


                                montoAfecto = costeo.Afecto;
                                montoIva = costeo.MontoIva;


                                montoExento = costeo.Exento;
                                
                                if (String.IsNullOrEmpty(montoExento))
                                {
                                    montoExento = "0";
                                }
                                montoTotal = (int.Parse(costeo.Afecto) + int.Parse(montoExento)+int.Parse(costeo.MontoIva)).ToString();
                                ajusteIva = costeo.AjusteIva;
                                facturaLeida.AjusteIva = costeo.AjusteIva;

                                facturaLeida.Glosa = costeo.Glosa;

                                if (ajusteIva == "")
                                {
                                    ajusteIva = facturaLeida.AjusteIva;
                                }

                                fechaDeDocumento = costeo.FechaDeDocumento;
                                fechaContableDelDocumento = costeo.FechaDeDocumento;
                                fechaDeVencimientoDelDocumento = costeo.FechaDeDocumento;


                            }

                        }

                     
                        if ((facturaSinG.TipoDeDocumento == "FACE" || facturaSinG.TipoDeDocumento == "NCCE") && facturaSinG.MontoExento != "" && facturaSinG.MontoExento != "0")
                        {
                            //es una factura con impuestos especiales

                        }
                        else if (facturaSinG.TipoDeDocumento == "FCEE")
                        {
                            //es una factura exenta
                        }

                        facturaSinG.TipoDeDocumento = tipoDeDocumento;
                        facturaSinG.NumeroDelDocumento = costeo.Folio;
                        facturaSinG.FechaDeDocumento = fechaDeDocumento;
                        facturaSinG.FechaContableDeDocumento = fechaContableDelDocumento;
                        facturaSinG.FechaDeVencimientoDeDocumento = fechaDeVencimientoDelDocumento;
                        facturaSinG.CodigoDeUnidadDeNegocio = codigoDeUnidadDeNegocio;
                        facturaSinG.RutCliente = costeo.Rut;
                        facturaSinG.DireccionDelCliente = direccionDelCliente;
                        facturaSinG.RutFacturador = "";
                        facturaSinG.CodigoVendedor = "";
                        facturaSinG.CodigoComisionista = "";
                        facturaSinG.Probabilidad = "";
                        facturaSinG.ListaPrecio = "";
                        facturaSinG.PlazoPago = "P01";
                        facturaSinG.MonedaDelDocumento = "CLP";
                        facturaSinG.TasaDeCambio = "";
                        facturaSinG.MontoAfecto = montoAfecto;
                        facturaSinG.MontoExento = montoExento;
                        facturaSinG.MontoIva = montoIva;
                        facturaSinG.MontoImpuestosEspecificos = "";
                        facturaSinG.MontoIvaRetenido = "";
                        facturaSinG.MontoImpuestosRetenidos = "";
                        facturaSinG.TipoDeDescuentoGlobal = "";
                        facturaSinG.DescuentoGlobal = "";
                        facturaSinG.TotalDelDocumento = montoTotal;
                        facturaSinG.DeudaPendiente = facturaSinG.TotalDelDocumento;
                        facturaSinG.TipoDocReferencia = "";
                        facturaSinG.NumDocReferencia = "";
                        facturaSinG.FechaDocReferencia = "";
                        facturaSinG.CodigoDelProducto = codigoDelProducto;
                        facturaSinG.Cantidad = "1";
                        facturaSinG.Unidad = "S.U.M";
                        facturaSinG.PrecioUnitario = precioUnitario;
                        facturaSinG.MonedaDelDetalle = "CLP";
                        facturaSinG.TasaDeCambio2 = "1";
                        facturaSinG.NumeroDeSerie = "";
                        facturaSinG.NumeroDeLote = "";
                        facturaSinG.FechaDeVencimiento = "";
                        facturaSinG.CentroDeCostos = costeo.CentroDeCosto;


                        switch (facturaSinG.CentroDeCostos)
                        {
                            case "ADMINISTRACION":
                                facturaSinG.CentroDeCostos = "203";
                                break;
                            case "INTERPLANTA":
                                facturaSinG.CentroDeCostos = "204";
                                facturaSinG.CodigoDeUnidadDeNegocio = "2";
                                break;
                            case "EMPRENDEDORES":
                                facturaSinG.CentroDeCostos = "208";
                                break;
                            case "ILLAPEL":
                                facturaSinG.CentroDeCostos = "205";
                                break;
                            case "SAN ANTONIO":
                                facturaSinG.CentroDeCostos = "207";
                                break;
                            case "MELIPILLA":
                                facturaSinG.CentroDeCostos = "200";
                                break;
                            case "SANTIAGO":
                                facturaSinG.CentroDeCostos = "206";
                                break;
                            case "RANCAGUA":
                                facturaSinG.CentroDeCostos = "201";
                                break;
                            case "CURICO":
                                facturaSinG.CentroDeCostos = "202";
                                break;
                            case "203":
                                facturaSinG.CentroDeCostos = "203";
                                break;
                            case "204":
                                facturaSinG.CentroDeCostos = "204";
                                facturaSinG.CodigoDeUnidadDeNegocio = "2";
                                break;
                            case "208":
                                facturaSinG.CentroDeCostos = "208";
                                break;
                            case "205":
                                facturaSinG.CentroDeCostos = "205";
                                break;
                            case "207":
                                facturaSinG.CentroDeCostos = "207";
                                break;
                            case "200":
                                facturaSinG.CentroDeCostos = "200";
                                break;
                            case "206":
                                facturaSinG.CentroDeCostos = "206";
                                break;
                            case "201":
                                facturaSinG.CentroDeCostos = "201";
                                break;
                            case "202":
                                facturaSinG.CentroDeCostos = "202";
                                break;
                            default:
                                facturaSinG.CentroDeCostos = "209";
                                break;

                        }

                        facturaSinG.TipoDeDescuento = "";
                        facturaSinG.Descuento = "";
                        facturaSinG.Ubicacion = "";
                        facturaSinG.Bodega = "";
                        facturaSinG.Concepto1 = "";
                        facturaSinG.Concepto2 = "";
                        facturaSinG.Concepto3 = "";
                        facturaSinG.Concepto4 = "";
                        facturaSinG.Descripcion = "";
                        facturaSinG.DescripcionAdicional = "";
                        facturaSinG.Stock = "0";
                        facturaSinG.Comentario11 = "";
                        facturaSinG.Comentario21 = "";
                        facturaSinG.Comentario31 = "";
                        facturaSinG.Comentario41 = "";
                        facturaSinG.Comentario51 = "";
                        facturaSinG.CodigoImpuestoEspecifico1 = "";
                        facturaSinG.MontoImpuestoEspecifico1 = "";
                        facturaSinG.CodigoImpuestoEspecifico2 = "";
                        facturaSinG.MontoImpuestoEspecifico2 = "";
                        facturaSinG.Modalidad = "";
                        facturaSinG.Glosa = costeo.Glosa;
                        facturaSinG.Referencia = "";
                        facturaSinG.FechaDeComprometida = "";
                        facturaSinG.PorcentajeCEEC = "";
                        facturaSinG.ImpuestoLey18211 = "";
                        facturaSinG.IvaLey18211 = "";
                        facturaSinG.CodigoKitFlexible = "";
                        facturaSinG.AjusteIva = ajusteIva;

                        //if (!String.IsNullOrEmpty(facturaSinG.RutCliente) && facturaSinG.RutCliente != "rut")
                        //{

                        //    if (int.Parse(facturaSinG.MontoExento)<0)
                        //    {
                        //        //Exento negativo,  se agrega fila extra
                        //        listadoDeFacturasCosteadas.Add(facturaConG);
                        //        listadoDeFacturasCosteadas.Add(facturaSinG);
                        //    }
                        //    else
                        //    {
                        //        //Exento positivo,  NO se agrega fila extra
                        //        listadoDeFacturasCosteadas.Add(facturaConG);
                        //    }

                        //}



                        facturaConG.TipoDeDocumento = tipoDeDocumento;
                        facturaConG.NumeroDelDocumento = costeo.Folio;
                        facturaConG.FechaDeDocumento = fechaDeDocumento;
                        facturaConG.FechaContableDeDocumento = fechaContableDelDocumento;
                        facturaConG.FechaDeVencimientoDeDocumento = fechaDeVencimientoDelDocumento;
                        facturaConG.CodigoDeUnidadDeNegocio = codigoDeUnidadDeNegocio;
                        facturaConG.RutCliente = costeo.Rut;
                        facturaConG.DireccionDelCliente = direccionDelCliente;
                        facturaConG.RutFacturador = "";
                        facturaConG.CodigoVendedor = "";
                        facturaConG.CodigoComisionista = "";
                        facturaConG.Probabilidad = "";
                        facturaConG.ListaPrecio = "";
                        facturaConG.PlazoPago = "P01";
                        facturaConG.MonedaDelDocumento = "CLP";
                        facturaConG.TasaDeCambio = "";
                        facturaConG.MontoAfecto = montoAfecto;
                        facturaConG.MontoExento = montoExento;
                        facturaConG.MontoIva = montoIva;
                        facturaConG.MontoImpuestosEspecificos = "";
                        facturaConG.MontoIvaRetenido = "";
                        facturaConG.MontoImpuestosRetenidos = "";
                        facturaConG.TipoDeDescuentoGlobal = "";
                        facturaConG.DescuentoGlobal = "";
                        facturaConG.TotalDelDocumento = montoTotal;
                        facturaConG.DeudaPendiente = facturaConG.TotalDelDocumento;
                        facturaConG.TipoDocReferencia = "";
                        facturaConG.NumDocReferencia = "";
                        facturaConG.FechaDocReferencia = "";
                        facturaConG.CodigoDelProducto = "410104G";
                        facturaConG.Cantidad = "1";
                        facturaConG.Unidad = "S.U.M";
                        facturaConG.PrecioUnitario = facturaConG.MontoExento;
                        facturaConG.MonedaDelDetalle = "CLP";
                        facturaConG.TasaDeCambio2 = "1";
                        facturaConG.NumeroDeSerie = "";
                        facturaConG.NumeroDeLote = "";
                        facturaConG.FechaDeVencimiento = "";
                        facturaConG.CentroDeCostos = costeo.CentroDeCosto;


                        switch (facturaConG.CentroDeCostos)
                        {
                            case "ADMINISTRACION":
                                facturaConG.CentroDeCostos = "203";
                                break;
                            case "INTERPLANTA":
                                facturaConG.CentroDeCostos = "204";
                                facturaConG.CodigoDeUnidadDeNegocio = "2";
                                break;
                            case "EMPRENDEDORES":
                                facturaConG.CentroDeCostos = "208";
                                break;
                            case "ILLAPEL":
                                facturaConG.CentroDeCostos = "205";
                                break;
                            case "SAN ANTONIO":
                                facturaConG.CentroDeCostos = "207";
                                break;
                            case "MELIPILLA":
                                facturaConG.CentroDeCostos = "200";
                                break;
                            case "SANTIAGO":
                                facturaConG.CentroDeCostos = "206";
                                break;
                            case "RANCAGUA":
                                facturaConG.CentroDeCostos = "201";
                                break;
                            case "CURICO":
                                facturaConG.CentroDeCostos = "202";
                                break;
                            case "203":
                                facturaConG.CentroDeCostos = "203";
                                break;
                            case "204":
                                facturaConG.CentroDeCostos = "204";
                                facturaConG.CodigoDeUnidadDeNegocio = "2";
                                break;
                            case "208":
                                facturaConG.CentroDeCostos = "208";
                                break;
                            case "205":
                                facturaConG.CentroDeCostos = "205";
                                break;
                            case "207":
                                facturaConG.CentroDeCostos = "207";
                                break;
                            case "200":
                                facturaConG.CentroDeCostos = "200";
                                break;
                            case "206":
                                facturaConG.CentroDeCostos = "206";
                                break;
                            case "201":
                                facturaConG.CentroDeCostos = "201";
                                break;
                            case "202":
                                facturaConG.CentroDeCostos = "202";
                                break;
                            default:
                                facturaConG.CentroDeCostos = "209";
                                break;

                        }

                        facturaConG.TipoDeDescuento = "";
                        facturaConG.Descuento = "";
                        facturaConG.Ubicacion = "";
                        facturaConG.Bodega = "";
                        facturaConG.Concepto1 = "";
                        facturaConG.Concepto2 = "";
                        facturaConG.Concepto3 = "";
                        facturaConG.Concepto4 = "";
                        facturaConG.Descripcion = "";
                        facturaConG.DescripcionAdicional = "";
                        facturaConG.Stock = "0";
                        facturaConG.Comentario11 = "";
                        facturaConG.Comentario21 = "";
                        facturaConG.Comentario31 = "";
                        facturaConG.Comentario41 = "";
                        facturaConG.Comentario51 = "";
                        facturaConG.CodigoImpuestoEspecifico1 = "";
                        facturaConG.MontoImpuestoEspecifico1 = "";
                        facturaConG.CodigoImpuestoEspecifico2 = "";
                        facturaConG.MontoImpuestoEspecifico2 = "";
                        facturaConG.Modalidad = "";
                        facturaConG.Glosa = costeo.Glosa;
                        facturaConG.Referencia = "";
                        facturaConG.FechaDeComprometida = "";
                        facturaConG.PorcentajeCEEC = "";
                        facturaConG.ImpuestoLey18211 = "";
                        facturaConG.IvaLey18211 = "";
                        facturaConG.CodigoKitFlexible = "";
                        facturaConG.AjusteIva = ajusteIva;

                        //if (!String.IsNullOrEmpty(facturaConG.RutCliente) && facturaConG.RutCliente != "rut")
                        //{
                        //    listadoDeFacturasCosteadas.Add(facturaConG);

                        //}


                        if (!String.IsNullOrEmpty(facturaSinG.RutCliente) && facturaSinG.RutCliente != "rut")
                        {

                            if (int.Parse(facturaSinG.MontoExento) < 0)
                            {
                                //Exento negativo,  se agrega fila extra
                                listadoDeFacturasCosteadas.Add(facturaSinG);
                                listadoDeFacturasCosteadas.Add(facturaConG);
                                
                               
                            }
                            else
                            {
                                //Exento positivo,  NO se agrega fila extra
                                listadoDeFacturasCosteadas.Add(facturaSinG);
                               
                            }

                        }

                    }

                }

            }

            //verificar coincidencias para agregar facturas que están costeadas, pero no presentes en Acepta
            foreach (var item in listadoDeCosteos)
            {
                Boolean coincidencia = false;
                foreach (var item2 in facturasLeidasEnPrimeraHoja)
                {
                    if (item.Folio==item2.NumeroDelDocumento && item.Rut==item2.RutCliente)
                    {
                        coincidencia = true;
                    }
                    
                }

                if (!coincidencia)
                {
                    //agregado el 22 de agosto del  2022, facturas ausentes en el listado de XML ahora pueden
                    //aparecer como costeadas si es que están en el listado de costeos.
                    //factura presente en costeo, pero no en listado de facturas xml
                    //crear nueva factura y agregarla a listado de costeadas

                    Factura fc = new Factura();

                    fc.TipoDeDocumento = "FACE";
                    fc.NumeroDelDocumento = item.Folio;
                    fc.FechaDeDocumento = item.FechaDeDocumento;
                    fc.FechaContableDeDocumento = item.FechaDeDocumento;
                    fc.FechaDeVencimientoDeDocumento = item.FechaDeDocumento;
                    fc.CodigoDeUnidadDeNegocio = "1";
                    fc.RutCliente = item.Rut;
                    fc.DireccionDelCliente = "Casa Matriz";
                    fc.RutFacturador = "";
                    fc.CodigoVendedor = "";
                    fc.CodigoComisionista = "";
                    fc.Probabilidad = "";
                    fc.ListaPrecio = "";
                    fc.PlazoPago = "P01";
                    fc.MonedaDelDocumento = "CLP";
                    fc.TasaDeCambio = "";
                    fc.MontoAfecto = item.Afecto;
                    fc.MontoExento = "0";
                    fc.MontoIva = item.MontoIva;
                    fc.MontoImpuestosEspecificos = "";
                    fc.MontoIvaRetenido = "";
                    fc.MontoImpuestosRetenidos = "";
                    fc.TipoDeDescuentoGlobal = "";
                    fc.DescuentoGlobal = "";
                    try
                    {
                        if (!String.IsNullOrEmpty(item.Afecto) && item.Afecto!="monto afecto")
                        {                         
                            fc.TotalDelDocumento = (int.Parse(fc.MontoAfecto) +int.Parse(fc.MontoExento)+ int.Parse(fc.MontoIva)).ToString();
                        }
                       
                    }
                    catch (Exception)
                    {
                        Console.WriteLine("hubo un error");
                        fc.TotalDelDocumento = "0";
                    }
                    
                    fc.DeudaPendiente = fc.TotalDelDocumento;
                    fc.TipoDocReferencia = "";
                    fc.NumDocReferencia = "";
                    fc.FechaDocReferencia = "";
                    fc.CodigoDelProducto = item.CodigoDelProducto;
                    fc.Cantidad = "1";
                    fc.Unidad = "S.U.M";
                    fc.PrecioUnitario = item.Afecto;

                    fc.MonedaDelDetalle = "CLP";
                    fc.TasaDeCambio2 = "1";
                    fc.NumeroDeSerie = "";
                    fc.NumeroDeLote = "";
                    fc.FechaDeVencimiento = "";
                    fc.CentroDeCostos = item.CentroDeCosto;


                    switch (fc.CentroDeCostos)
                    {
                        case "ADMINISTRACION":
                            fc.CentroDeCostos = "203";
                            break;
                        case "INTERPLANTA":
                            fc.CentroDeCostos = "204";
                            fc.CodigoDeUnidadDeNegocio = "2";
                            break;
                        case "EMPRENDEDORES":
                            fc.CentroDeCostos = "208";
                            break;
                        case "ILLAPEL":
                            fc.CentroDeCostos = "205";
                            break;
                        case "SAN ANTONIO":
                            fc.CentroDeCostos = "207";
                            break;
                        case "MELIPILLA":
                            fc.CentroDeCostos = "200";
                            break;
                        case "SANTIAGO":
                            fc.CentroDeCostos = "206";
                            break;
                        case "RANCAGUA":
                            fc.CentroDeCostos = "201";
                            break;
                        case "CURICO":
                            fc.CentroDeCostos = "202";
                            break;
                        case "203":
                            fc.CentroDeCostos = "203";
                            break;
                        case "204":
                            fc.CentroDeCostos = "204";
                            fc.CodigoDeUnidadDeNegocio = "2";
                            break;
                        case "208":
                            fc.CentroDeCostos = "208";
                            break;
                        case "205":
                            fc.CentroDeCostos = "205";
                            break;
                        case "207":
                            fc.CentroDeCostos = "207";
                            break;
                        case "200":
                            fc.CentroDeCostos = "200";
                            break;
                        case "206":
                            fc.CentroDeCostos = "206";
                            break;
                        case "201":
                            fc.CentroDeCostos = "201";
                            break;
                        case "202":
                            fc.CentroDeCostos = "202";
                            break;
                        default:
                            fc.CentroDeCostos = "209";
                            break;
                    }


                    fc.TipoDeDescuento = "";
                    fc.Descuento = "";
                    fc.Ubicacion = "";
                    fc.Bodega = "";
                    fc.Concepto1 = "";
                    fc.Concepto2 = "";
                    fc.Concepto3 = "";
                    fc.Concepto4 = "";
                    fc.Descripcion = "";
                    fc.DescripcionAdicional = "";
                    fc.Stock = "0";
                    fc.Comentario11 = "";
                    fc.Comentario21 = "";
                    fc.Comentario31 = "";
                    fc.Comentario41 = "";
                    fc.Comentario51 = "";
                    fc.CodigoImpuestoEspecifico1 = "";
                    fc.MontoImpuestoEspecifico1 = "";
                    fc.CodigoImpuestoEspecifico2 = "";
                    fc.MontoImpuestoEspecifico2 = "";
                    fc.Modalidad = "";
                    fc.Glosa = "Factura ausente en Acepta";
                    fc.Referencia = "";
                    fc.FechaDeComprometida = "";
                    fc.PorcentajeCEEC = "";
                    fc.ImpuestoLey18211 = "";
                    fc.IvaLey18211 = "";
                    fc.CodigoKitFlexible = "";
                    fc.AjusteIva = item.AjusteIva;

                    if (!String.IsNullOrEmpty(fc.RutCliente) && fc.RutCliente != "rut")
                    {
                        listadoDeFacturasCosteadas.Add(fc);

                        if(fc.CodigoDelProducto == "410104")
                        {
                            Factura facturaPetroleo = new Factura();


                            facturaPetroleo.TipoDeDocumento = fc.TipoDeDocumento;
                            facturaPetroleo.NumeroDelDocumento = fc.NumeroDelDocumento;
                            facturaPetroleo.FechaDeDocumento = fc.FechaDeDocumento;
                            facturaPetroleo.FechaContableDeDocumento = fc.FechaContableDeDocumento;
                            facturaPetroleo.FechaDeVencimientoDeDocumento = fc.FechaDeVencimientoDeDocumento;
                            facturaPetroleo.CodigoDeUnidadDeNegocio = fc.CodigoDeUnidadDeNegocio;
                            facturaPetroleo.RutCliente = fc.RutCliente;
                            facturaPetroleo.DireccionDelCliente = fc.DireccionDelCliente;
                            facturaPetroleo.RutFacturador = fc.RutFacturador;
                            facturaPetroleo.CodigoVendedor = fc.CodigoVendedor;
                            facturaPetroleo.CodigoComisionista = fc.CodigoComisionista;
                            facturaPetroleo.Probabilidad = fc.Probabilidad;
                            facturaPetroleo.ListaPrecio = fc.ListaPrecio;
                            facturaPetroleo.PlazoPago = fc.PlazoPago;
                            facturaPetroleo.MonedaDelDocumento = fc.MonedaDelDocumento;
                            facturaPetroleo.TasaDeCambio = fc.TasaDeCambio;
                            facturaPetroleo.MontoAfecto = fc.MontoAfecto;
                            facturaPetroleo.MontoExento = fc.MontoExento;
                            facturaPetroleo.MontoIva = fc.MontoIva;
                            facturaPetroleo.MontoImpuestosEspecificos = fc.MontoImpuestosEspecificos;
                            facturaPetroleo.MontoIvaRetenido = fc.MontoIvaRetenido;
                            facturaPetroleo.MontoImpuestosRetenidos = fc.MontoImpuestosRetenidos;
                            facturaPetroleo.TipoDeDescuentoGlobal = fc.TipoDeDescuentoGlobal;
                            facturaPetroleo.DescuentoGlobal = fc.DescuentoGlobal;
                            facturaPetroleo.TotalDelDocumento = fc.TotalDelDocumento;
                            facturaPetroleo.DeudaPendiente = fc.DeudaPendiente;
                            facturaPetroleo.TipoDocReferencia = fc.TipoDocReferencia;
                            facturaPetroleo.NumDocReferencia = fc.NumDocReferencia;
                            facturaPetroleo.FechaDocReferencia = fc.FechaDocReferencia;
                            facturaPetroleo.CodigoDelProducto = fc.CodigoDelProducto;
                            facturaPetroleo.Cantidad = fc.Cantidad;
                            facturaPetroleo.Unidad = fc.Unidad;
                            facturaPetroleo.PrecioUnitario = fc.PrecioUnitario;
                            facturaPetroleo.MonedaDelDetalle = fc.MonedaDelDetalle;
                            facturaPetroleo.TasaDeCambio2 = fc.TasaDeCambio2;
                            facturaPetroleo.NumeroDeSerie = fc.NumeroDeSerie;
                            facturaPetroleo.NumeroDeLote = fc.NumeroDeLote;
                            facturaPetroleo.FechaDeVencimiento = fc.FechaDeVencimiento;
                            facturaPetroleo.CentroDeCostos = fc.CentroDeCostos;
                            facturaPetroleo.TipoDeDescuento = fc.TipoDeDescuento;
                            facturaPetroleo.Descuento = fc.Descuento;
                            facturaPetroleo.Ubicacion = fc.Ubicacion;
                            facturaPetroleo.Bodega = fc.Bodega;
                            facturaPetroleo.Concepto1 = fc.Concepto1;
                            facturaPetroleo.Concepto2 = fc.Concepto2;
                            facturaPetroleo.Concepto3 = fc.Concepto3;
                            facturaPetroleo.Concepto4 = fc.Concepto4;
                            facturaPetroleo.Descripcion = fc.Descripcion;
                            facturaPetroleo.DescripcionAdicional = fc.DescripcionAdicional;
                            facturaPetroleo.Stock = fc.Stock;
                            facturaPetroleo.Comentario11 = fc.Comentario11;
                            facturaPetroleo.Comentario21 = fc.Comentario21;
                            facturaPetroleo.Comentario31 = fc.Comentario31;
                            facturaPetroleo.Comentario41 = fc.Comentario41;
                            facturaPetroleo.Comentario51 = fc.Comentario51;
                            facturaPetroleo.CodigoImpuestoEspecifico1 = fc.CodigoImpuestoEspecifico1;
                            facturaPetroleo.MontoImpuestoEspecifico1 = fc.MontoImpuestoEspecifico1;
                            facturaPetroleo.CodigoImpuestoEspecifico2 = fc.CodigoImpuestoEspecifico2;
                            facturaPetroleo.MontoImpuestoEspecifico2 = fc.MontoImpuestoEspecifico2;
                            facturaPetroleo.Modalidad = fc.Modalidad;
                            facturaPetroleo.Glosa = fc.Glosa;
                            facturaPetroleo.Referencia = fc.Referencia;
                            facturaPetroleo.FechaDeComprometida = fc.FechaDeComprometida;
                            facturaPetroleo.PorcentajeCEEC = fc.PorcentajeCEEC;
                            facturaPetroleo.ImpuestoLey18211 = fc.ImpuestoLey18211;
                            facturaPetroleo.IvaLey18211 = fc.IvaLey18211;
                            facturaPetroleo.CodigoKitFlexible = fc.CodigoKitFlexible;
                            facturaPetroleo.AjusteIva = fc.AjusteIva;
                            facturaPetroleo.CodigoDelProducto = "410104G";
                            facturaPetroleo.PrecioUnitario = fc.MontoExento;

                           
                            listadoDeFacturasCosteadas.Add(facturaPetroleo);
                        }
                    }
                    
                }
            }


            return listadoDeFacturasCosteadas;

        }

        private String convertirFechaDePamelaAFechaParaExcelDeManager(String fechaDePamela)
        {

            String fechaConvertida = "";

            Char espacio = ' ';

            try
            {
                if (fechaDePamela.Contains(espacio))
                {
                    //viene como fecha

                    string[] partes = fechaDePamela.Split(' ');

                    string[] partesDeFecha = partes[0].Split('/');

                    String fechaQueDebieseSer = partesDeFecha[1] + "/" + partesDeFecha[0] + "/20" + partesDeFecha[2];

                    fechaConvertida = fechaQueDebieseSer;

                }
                else
                {
                    //no viene como fecha
                    fechaConvertida = fechaDePamela;
                }
            }
            catch (Exception)
            {

              
            }


            return fechaConvertida;
        }


    }


}
