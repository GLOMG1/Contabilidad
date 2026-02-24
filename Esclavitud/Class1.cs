using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.WindowsAPICodePack.Dialogs;
using Amazon.Lambda.Model;
using System.Xml.Linq;
using System.IO;
using System.Linq;
using Amazon.SimpleEmail;
using System.Windows.Controls;
using System.Threading;
using System.Security.Cryptography.Xml;

namespace Esclavitud
{
    internal class Class1
    {
        public Excel.Application archivo { get; }
        public Excel.Workbook libro { get; }
        public Excel.Worksheet hoja { get; }

        public Class1()
        {
            archivo = Globals.ThisAddIn.Application;
            libro = archivo.ActiveWorkbook;
            hoja = archivo.ActiveSheet;
        }
        public void crearTablas()
        {
            string[] nombreHojas = { "Egresos", "Ingresos", "Bancos" };
            object[,] tblEgresos = new object[1, 47]
            {
                {"ID","Fiscal","Validación","Serie y Folio","RFC","Nombre","Fecha","Tipo","Uso", "Version", "UUID","UUID Relacionado NC","UUID Relacionado REP", "Moneda", "Forma de Pago", "Medoto de Pago", "Condicion de Pago","Concepto", "Centro de Costos", "Gravado 16","Gravado 8", "Gravado 0","Excento", "Sin IVA", "Descuento 16", "Descuento 8", "Descuento 0", "Descuento Excento", "Descuento Sin IVA","SubTotal","IVA","ISH","ISR Retenido","IVA Retenido","Total","Porcion de Pago", "Fecha de Pago", "Fecha de Pago Efectiva","Porciento ISR","Porciento IVA", "Estado BD","ID Registro Banco", "Fecha Bancos","Folio Bancos","Estatus Banco", "Banco", "Numero Cuenta"}
            };

            string[] formatoC = new string[47]
            {
                "#,##0",    // ID
                "@",        // Fiscal
                "@",        // Validación
                "@",        // Serie y Folio
                "@",        // RFC
                "@",        // Nombre
                "dd/MM/yyyy", // Fecha
                "@",        // Tipo
                "@",        // Uso
                "@",        // Version
                "@",        // UUID
                "@",        // UUID Relacionado NC
                "@",        // UUID Relacionado REP
                "@",        // Moneda
                "@",        // Forma de Pago
                "@",        // Metodo de Pago
                "@",        // Condicion de Pago
                "@",        // Concepto
                "@",        // Centro de Costos
                "#,##0.00", // Gravado 16
                "#,##0.00", // Gravado 8
                "#,##0.00", // Gravado 0
                "#,##0.00", // Excento
                "#,##0.00", // Sin IVA
                "#,##0.00", // Descuento 16
                "#,##0.00", // Descuento 8
                "#,##0.00", // Descuento 0
                "#,##0.00", // Descuento Excento
                "#,##0.00", // Descuento Sin IVA
                "#,##0.00", // SubTotal
                "#,##0.00", // IVA
                "#,##0.00", // ISH
                "#,##0.00", // ISR Retenido
                "#,##0.00", // IVA Retenido
                "#,##0.00", // Total
                "#,##0.00", // Porcion de Pago
                "dd/MM/yyyy", // Fecha de Pago
                "dd/MM/yyyy", // Fecha de Pago Efectiva
                "#,##0.00", // Porciento ISR
                "#,##0.00", // Porciento IVA
                "@",        // Estado BD
                "@",        // ID Registro Banco
                "@",        // Fecha Bancos
                "@",        // Folio Bancos
                "@",        // Estatus Banco
                "@",        // Banco
                "@",        // Numero Cuenta
};

            object[,] tblBancos = new object[1,7]
            {
                {"ID", "Fecha Operacion", "Concepto", "Referencia", "Retiros", "Depositos", "Saldo Liquidacion"}
            };

            foreach (string x in nombreHojas) 
            {
                Excel.Worksheet nuevaHoja = libro.Sheets.Add(After: libro.Sheets[libro.Sheets.Count]);
                nuevaHoja.Name = x;

                switch (x)
                {
                    case "Egresos":
                        nuevaHoja.Range["B4:AV4"].Value = tblEgresos;

                        Excel.ListObject egresos = nuevaHoja.ListObjects.Add(
                            Excel.XlListObjectSourceType.xlSrcRange,
                            nuevaHoja.Range["B4:AV4"],
                            null,
                            Excel.XlYesNoGuess.xlYes
                            );
                        egresos.TableStyle = "TableStyleLight21";
                        egresos.Name = "tblEgresos";
                        egresos.ShowAutoFilter = false;
                        for (int cont = 1; cont <= 22; cont++)
                        {
                            var celda = egresos.ListColumns[cont].Range;
                            celda.NumberFormat = formatoC[cont-1];
                        }
                        
                        break;
                    case "Ingresos":
                        nuevaHoja.Range["B4:W4"].Value = tblEgresos;

                        Excel.ListObject ingresos = nuevaHoja.ListObjects.Add(
                            Excel.XlListObjectSourceType.xlSrcRange,
                            nuevaHoja.Range["B4:W4"],
                            null,
                            Excel.XlYesNoGuess.xlYes
                            );
                        ingresos.TableStyle = "TableStyleLight21";
                        ingresos.Name = "tblIngresos";
                        ingresos.ShowAutoFilter = false;
                        for (int cont = 1; cont <= 22; cont++)
                        {
                            var celda = ingresos.ListColumns[cont].Range;
                            celda.NumberFormat = formatoC[cont - 1];
                        }

                        break;
                    case "Bancos":
                        nuevaHoja.Range["B4:H4"].Value = tblBancos;

                        Excel.ListObject bancos = nuevaHoja.ListObjects.Add(
                            Excel.XlListObjectSourceType.xlSrcRange,
                            nuevaHoja.Range["B4:H4"],
                            null,
                            Excel.XlYesNoGuess.xlYes
                            );
                        bancos.TableStyle = "TableStyleLight21";
                        bancos.Name = "tblBancos";
                        bancos.ShowAutoFilter = false;

                        break;
                }
            }
        }
        public string Direccion()
        {
            CommonOpenFileDialog dialogo = new CommonOpenFileDialog()
            {
                InitialDirectory = "C:\\Users\\Jaime\\Documents\\Python\\Lector de CFDIs\\Egresos\\01- Ene",
                IsFolderPicker = true
            };
            if (dialogo.ShowDialog() == CommonFileDialogResult.Ok)
            {
                // Obtener la ruta seleccionada
                string nombreRuta = dialogo.FileName;
                return nombreRuta;
            }
            else
            {
                return null;
            }
        }
        void barraProgreso(Excel.Application excelApp, int actual, int total, string operacion = "Procesando")
        {
            excelApp.StatusBar = $"{operacion} {actual} de {total}";
        }
        public void importar()
        {
            string carpeta = Direccion();
            string[] cfdi = Directory.GetFiles(carpeta, "*.xml");

            Excel.Worksheet hojaEgresos = libro.Sheets["Egresos"];
            Excel.ListObject tabla = hojaEgresos.ListObjects["tblEgresos"];

            DialogResult verDetalle = MessageBox.Show("¿Quieres separar los conceptos de tu factura?", "Conceptos", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            archivo.ScreenUpdating = false;
            archivo.DisplayAlerts = false;
            archivo.EnableEvents = false;
            archivo.Calculation = Excel.XlCalculation.xlCalculationManual;

            int n = cfdi.Count();
            int p = 1;

            foreach (string x in cfdi) 
            {
                barraProgreso(archivo, p, n);
                p++;
                XDocument doc = XDocument.Load(x);
                
                /// -------------------------------------------------------------------------------------
                string[,] info = lectorPorTipos(doc, verDetalle);
                /// -------------------------------------------------------------------------------------
                
                for (int i = 0; i <= (info.GetLength(0)-1); i++) {
                    Excel.ListRow añadir = tabla.ListRows.Add();
                    for (int y = 1; y <= 21; y++) {
                        añadir.Range[1, y].Value2 = info[i, y];
                    }
                }
            }
            /// Aqui inicia el analisis de datos :o -----------------------------------------------------
            /// 1ro Verificar filas vacias
            for (int i = tabla.ListRows.Count; i >= 1; i--)
            {
                Excel.Range rowRange = tabla.ListRows[i].Range;
                if (string.IsNullOrWhiteSpace(rowRange.Cells[1, 8].Text.ToString()))
                {
                    tabla.ListRows[i].Delete();
                }
            }
            /// -------------------------------------------------------------------------------------
            tabla = hojaEgresos.ListObjects["tblEgresos"];
            analisisDatos.REPs(tabla);
            /// -------------------------------------------------------------------------------------
            ///2do Poner el ID
            for (int i = tabla.ListRows.Count; i >= 1; i--)
            {
                Excel.Range rowRange = tabla.ListRows[i].Range;
                if (string.IsNullOrWhiteSpace(rowRange.Cells[1, 1].Text.ToString()))
                {
                    rowRange.Cells[1, 1].Value = i;
                }
            }
            

            archivo.StatusBar = false;
            archivo.ScreenUpdating = true;
            archivo.DisplayAlerts = true;
            archivo.EnableEvents = true;
            archivo.Calculation = Excel.XlCalculation.xlCalculationAutomatic;

            // Forzar actualización
            archivo.ActiveWindow.View = Excel.XlWindowView.xlNormalView;
            archivo.ActiveSheet.Calculate();
        }

        public string[,] lectorPorTipos(XDocument doc, DialogResult verDetalle)
        {
            /// Declaracion de variables a utilizar -----------------------------------------------------
            XNamespace cfdi = "http://www.sat.gob.mx/cfd/4";
            XNamespace pago20 = "http://www.sat.gob.mx/Pagos20";
            XNamespace tfd = "http://www.sat.gob.mx/TimbreFiscalDigital";

            var comprobante = doc.Element(cfdi + "Comprobante");

            string[,] informacionCFDI = null;
            var UUID = comprobante.Attribute("TipoDeComprobante").Value;

            switch (UUID)
            {
                case "I":
                    {
                        if (verDetalle == DialogResult.Yes)
                        {
                            var conceptos = comprobante.Element(cfdi + "Conceptos").Elements(cfdi + "Concepto");
                            int numeroConceptos = conceptos.Count();
                            int cont = 0;
                            informacionCFDI = new string[numeroConceptos, 22];

                            foreach(var concepto in conceptos)
                            {
                                /// INFORMACION GENERAL
                                informacionCFDI[cont, 2] = comprobante.Attribute("Serie")?.Value + " - " + comprobante.Attribute("Folio")?.Value;
                                informacionCFDI[cont, 3] = comprobante.Element(cfdi + "Emisor").Attribute("Rfc")?.Value;
                                informacionCFDI[cont, 4] = comprobante.Element(cfdi + "Emisor").Attribute("Nombre")?.Value;
                                informacionCFDI[cont, 5] = comprobante.Attribute("Fecha")?.Value;
                                informacionCFDI[cont, 6] = comprobante.Attribute("TipoDeComprobante")?.Value;
                                informacionCFDI[cont, 7] = comprobante.Element(cfdi + "Receptor").Attribute("UsoCFDI")?.Value;
                                informacionCFDI[cont, 8] = comprobante.Element(cfdi + "Complemento").Element(tfd + "TimbreFiscalDigital").Attribute("UUID")?.Value;
                                    //informacionCFDI[cont, 9] = No puede tener doc relacionados 
                                informacionCFDI[cont, 10] = comprobante.Attribute("Moneda")?.Value;
                                informacionCFDI[cont, 11] = comprobante.Attribute("FormaPago")?.Value;
                                informacionCFDI[cont, 12] = comprobante.Attribute("MetodoPago")?.Value;
                                informacionCFDI[cont, 13] = comprobante.Attribute("CondicionesDePago")?.Value;
                                informacionCFDI[cont, 14] = concepto.Attribute("Descripcion")?.Value;

                                /// INFORMACION DE IMPUESTOS
                                string objetoImpuesto = concepto.Attribute("ObjetoImp")?.Value;
                                if (objetoImpuesto == "02")
                                {
                                    informacionCFDI[cont, 15] = concepto.Element(cfdi + "Impuestos").Element(cfdi + "Traslados").Element(cfdi + "Traslado").Attribute("Impuesto").Value;
                                    informacionCFDI[cont, 16] = concepto.Element(cfdi + "Impuestos").Element(cfdi + "Traslados").Element(cfdi + "Traslado").Attribute("TasaOCuota").Value;
                                    switch (informacionCFDI[cont, 15])
                                    {
                                        case "002": ///IVA
                                            switch (informacionCFDI[cont, 16])
                                            {
                                                case "0.160000":
                                                    informacionCFDI[cont, 15] = concepto.Element(cfdi + "Impuestos").Element(cfdi + "Traslados").Element(cfdi + "Traslado").Attribute("Importe").Value;
                                                    informacionCFDI[cont, 16] = "";
                                                    break;
                                                case "0.000000":
                                                    informacionCFDI[cont, 16] = concepto.Element(cfdi + "Impuestos").Element(cfdi + "Traslados").Element(cfdi + "Traslado").Attribute("Importe").Value;
                                                    informacionCFDI[cont, 15] = "";
                                                    break;
                                                default:
                                                    MessageBox.Show("Nuevo Tipo de taza del IVA desbloqueada:" + informacionCFDI[cont, 16]);
                                                    break;
                                            }
                                            break;
                                        default:
                                            MessageBox.Show("Nuevo Tipo de impuesto desbloqueado:" + informacionCFDI[cont, 15]);
                                            break;
                                    }
                                }
                                else
                                {
                                    informacionCFDI[cont, 16] = concepto.Attribute("Importe")?.Value;
                                }
                                /// TOTALES Y COMPROBACIONES
                                informacionCFDI[cont, 17] = concepto.Attribute("Importe")?.Value;
                                informacionCFDI[cont, 18] = informacionCFDI[cont, 15];
                                informacionCFDI[cont, 21] = "Putos";

                                cont ++;
                            }
                        }
                        else
                        {
                            var conceptos = comprobante.Element(cfdi + "Conceptos").Elements(cfdi + "Concepto");
                            informacionCFDI = new string[1, 22];

                            /// INFORMACION GENERAL
                            informacionCFDI[0, 2] = comprobante.Attribute("Serie")?.Value + " - " + comprobante.Attribute("Folio")?.Value;
                            informacionCFDI[0, 3] = comprobante.Element(cfdi + "Emisor").Attribute("Rfc")?.Value;
                            informacionCFDI[0, 4] = comprobante.Element(cfdi + "Emisor").Attribute("Nombre")?.Value;
                            informacionCFDI[0, 5] = comprobante.Attribute("Fecha")?.Value;
                            informacionCFDI[0, 6] = comprobante.Attribute("TipoDeComprobante")?.Value;
                            informacionCFDI[0, 7] = comprobante.Element(cfdi + "Receptor").Attribute("UsoCFDI")?.Value;
                            informacionCFDI[0, 8] = comprobante.Element(cfdi + "Complemento").Element(tfd + "TimbreFiscalDigital").Attribute("UUID")?.Value;
                            //informacionCFDI[cont, 9] = No puede tener doc relacionados 
                            informacionCFDI[0, 10] = comprobante.Attribute("Moneda")?.Value;
                            informacionCFDI[0, 11] = comprobante.Attribute("FormaPago")?.Value;
                            informacionCFDI[0, 12] = comprobante.Attribute("MetodoPago")?.Value;
                            informacionCFDI[0, 13] = comprobante.Attribute("CondicionesDePago")?.Value;
                            foreach (var concepto in conceptos)
                            {
                                informacionCFDI[0, 14] = concepto.Attribute("Descripcion")?.Value + " | " + informacionCFDI[0, 14];
                            }
                            informacionCFDI[0, 17] = comprobante.Attribute("SubTotal")?.Value;
                            informacionCFDI[0, 21] = comprobante.Attribute("Total")?.Value;
                        }
                }
                    break;
                case "P":
                {
                        var documentosRelacionados = comprobante.Element(cfdi + "Complemento").Element(pago20 + "Pagos").Element(pago20 + "Pago").Elements(pago20 + "DoctoRelacionado");
                        int numeroDocsR = documentosRelacionados.Count();
                        int cont = 0;

                        informacionCFDI = new string[numeroDocsR, 22];

                        foreach (var docRelacionado in documentosRelacionados)
                        {
                            MessageBox.Show(docRelacionado.ToString());
                            informacionCFDI[cont, 2] = comprobante.Attribute("Serie")?.Value + " - " + comprobante.Attribute("Folio")?.Value;
                            informacionCFDI[cont, 3] = comprobante.Element(cfdi + "Emisor").Attribute("Rfc")?.Value;
                            informacionCFDI[cont, 4] = comprobante.Element(cfdi + "Emisor").Attribute("Nombre")?.Value;
                            informacionCFDI[cont, 5] = comprobante.Attribute("Fecha")?.Value;
                            informacionCFDI[cont, 6] = comprobante.Attribute("TipoDeComprobante")?.Value;
                            informacionCFDI[cont, 8] = comprobante.Element(cfdi + "Complemento").Element(tfd + "TimbreFiscalDigital").Attribute("UUID")?.Value;
                            informacionCFDI[cont, 9] = docRelacionado.Attribute("IdDocumento")?.Value;
                            informacionCFDI[cont, 21] = docRelacionado.Attribute("ImpPagado")?.Value;
                            cont++;
                        }
                }
                    break;
                case "E":

                    break;
                case "N":

                    break;
                default:
                    informacionCFDI = new string[0, 22];
                    break;
            }
            return informacionCFDI ?? new string[0, 22];
        }
    }


}
