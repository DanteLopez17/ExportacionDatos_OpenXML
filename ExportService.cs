using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;

namespace ExportacionDatos_OpenXML
{
    public class ExportService
    {
        //crear un archivo de Excel en la ruta dada.
        public void CrearExcel(TestModelList data, string OutPutFileDirectory)
        {
            var datetime = DateTime.Now.ToString().Replace("/", "_").Replace(":", "_");

            string fileFullname = Path.Combine(OutPutFileDirectory, "Output.xlsx");

            if (File.Exists(fileFullname))
            {
                fileFullname = Path.Combine(OutPutFileDirectory, "Output_" + datetime + ".xlsx");
            }

            using (SpreadsheetDocument package = SpreadsheetDocument.Create(fileFullname, SpreadsheetDocumentType.Workbook))
            {
                CrearPartesExcel(package, data);
            }
        }

        //crear un libro de trabajo y una hoja de trabajo en Excel.
        private void CrearPartesExcel(SpreadsheetDocument document, TestModelList data)
        {
            //Hoja de datos
            SheetData partSheetData = GenerarHojadeDatos(data);

            //Libro de trabajo
            WorkbookPart workbookPart1 = document.AddWorkbookPart();
            GenerarParteContenidoLibro(workbookPart1);

            //Estilos libro
            WorkbookStylesPart workbookStylesPart1 = workbookPart1.AddNewPart<WorkbookStylesPart>("rId3");
            GenerarEstilosLibro(workbookStylesPart1);

            //Estilos
            //WorkbookStylesPart wbsp = workbookPart1.AddNewPart<WorkbookStylesPart>();
            //wbsp.Stylesheet = CreateStylesheet();
            //wbsp.Stylesheet.Save();

            //Hoja de calculo
            WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("rId1");
            GenerarContenidoHojaCalculo(worksheetPart1, partSheetData);
        }

        //crear contenido de libros, recibe por paramaetro parte de libro de trabajo
        private void GenerarParteContenidoLibro(WorkbookPart workbookPart1)
        {
            //Libro de trabajo (archivo completo)
            Workbook workbook1 = new Workbook();
            //Contenedor de hojas
            Sheets sheets1 = new Sheets();
            //Hoja individual
            Sheet sheet1 = new Sheet() { Name = "Hoja 1", SheetId = (UInt32Value)1U, Id = "rId1"};

            //Crea una hoja mas, algo tiene el Id que lo rompe, crea la hoja vacia pero lanza un error al abrir el excel
            //Sheet sheet2 = new Sheet() { Name = "Hoja 2", SheetId = (UInt32Value)2U, Id = "rId2" };


            //Agrega las hoja creada
            sheets1.Append(sheet1);

            //sheets1.Append(sheet2);

            workbook1.Append(sheets1);
            workbookPart1.Workbook = workbook1;
        }

        //crear contenido de hojas de calculo en Excel
        private void GenerarContenidoHojaCalculo(WorksheetPart worksheetPart1, SheetData sheetData1)
        {
            //Hoja de calculo individual
            Worksheet worksheet1 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            //Dimension de hoja
            SheetDimension sheetDimension1 = new SheetDimension() { Reference = "A1" };

            //Contenedor de Vistas
            SheetViews sheetViews1 = new SheetViews();
            //Vista de hoja
            SheetView sheetView1 = new SheetView() { TabSelected = true, WorkbookViewId = (UInt32Value)0U };
            //Celda seleccionada en la apertura del archivo, deben coincidir o da error 
            Selection selection1 = new Selection() { ActiveCell = "A1", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "A1" } };

            //Adjunta la seleccion de la celda a la vista de la hoja
            sheetView1.Append(selection1);
            //Adjunta la hoja al contenedor de vistas
            sheetViews1.Append(sheetView1);

            //FormatoPropiedades = DefaultRowHeight(altura de la fila)
            SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };

            PageMargins pageMargins1 = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            worksheet1.Append(sheetDimension1);
            worksheet1.Append(sheetViews1);
            worksheet1.Append(sheetFormatProperties1);
            worksheet1.Append(sheetData1);
            worksheet1.Append(pageMargins1);
            worksheetPart1.Worksheet = worksheet1;
        }

        //estilos de libros de trabajo proporcionando su propio tamaño de fuente, color, nombre de fuente, propiedades de borde, formatos de estilo de celda, etc.
        private void GenerarEstilosLibro(WorkbookStylesPart workbookStylesPart1)
        {
            //Estilo de hoja
            Stylesheet stylesheet1 = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            #region Font
            // ------------------ Contenedor de fuentes
            Fonts fonts1 = new Fonts() { Count = (UInt32Value)2U};

            //Fuente 1 Individual (aparentemente sobre la informacion, no incluye los encabezados)
            Font font1 = new Font();
            //Caracteristicas para la fuente 1
            FontSize fontSize1 = new FontSize() { Val = 11D };
            //No cambia el color
            Color color1 = new Color() { Rgb = new HexBinaryValue("FF3333") };
            //No cambia la fuente
            FontName fontName1 = new FontName() { Val = "Arial" };
            //FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
            //FontScheme fontScheme1 = new FontScheme() { Val = FontSchemeValues.Minor };

            //Adjuntan las caracteristicas a la fuente 1
            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            //font1.Append(fontFamilyNumbering1);
            //font1.Append(fontScheme1);

            //Fuente 2 Individual (Sobre los elementos que tienen 2U)
            Font font2 = new Font();
            //Bold bold1 = new Bold();
            //Caracteristicas de la fuente 2
            FontSize fontSize2 = new FontSize() { Val = 12D };
            //No cambia el color
            Color color2 = new Color() { Rgb = new HexBinaryValue("2b00ff")  };
            //No cambia la fuente
            FontName fontName2 = new FontName() { Val = "Calibri" };
            //FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 2 };
            //FontScheme fontScheme2 = new FontScheme() { Val = FontSchemeValues.Minor };

            //Adjunta las caracteristicas a la fuente 2
            //font2.Append(bold1);
            font2.Append(fontSize2);
            font2.Append(color2);
            font2.Append(fontName2);
            //font2.Append(fontFamilyNumbering2);
            //font2.Append(fontScheme2);

            //Adjunta las fuentes al contenedor de fuentes
            fonts1.Append(font1);
            fonts1.Append(font2);

            #endregion

            #region Fill
            Fills fills1 = new Fills() { Count = (UInt32Value)0U };

            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };

            fill1.Append(patternFill1);

            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };

            fill2.Append(patternFill2);

            // Fill ID = 2
            Fill fill3 = new Fill();
            PatternFill patternFill3 = new PatternFill() { PatternType = PatternValues.Solid };    // customized Pattern value
            ForegroundColor foregroundColorFill3 = new ForegroundColor() { Rgb = new HexBinaryValue("FF00FF00") };
            BackgroundColor backgroundColorFill3 = new BackgroundColor() { Rgb = new HexBinaryValue("5B9BD500") };

            patternFill3.Append(foregroundColorFill3);
            patternFill3.Append(backgroundColorFill3);
            fill3.Append(patternFill3);

            // Fill ID = 3
            Fill fill4 = new Fill();
            PatternFill patternFill4 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColorFill4 = new ForegroundColor() { Rgb = new HexBinaryValue("56EFD8") };
            BackgroundColor backgroundColorFill4 = new BackgroundColor() { Rgb = new HexBinaryValue("FF00FF00") };

            patternFill4.Append(foregroundColorFill4);
            patternFill4.Append(backgroundColorFill4);
            fill4.Append(patternFill4);

            fills1.Append(fill1);
            fills1.Append(fill2);
            fills1.Append(fill3);
            fills1.Append(fill4);

            #endregion

            #region Border
            Borders borders1 = new Borders() { Count = (UInt32Value)2U };

            Border border1 = new Border();
            LeftBorder leftBorder1 = new LeftBorder();
            RightBorder rightBorder1 = new RightBorder();
            TopBorder topBorder1 = new TopBorder();
            BottomBorder bottomBorder1 = new BottomBorder();
            DiagonalBorder diagonalBorder1 = new DiagonalBorder();

            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);

            //Esta tomando este borde, y es el borde de la tabla
            Border border2 = new Border();

            LeftBorder leftBorder2 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color3 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder2.Append(color3);

            RightBorder rightBorder2 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color4 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder2.Append(color4);

            TopBorder topBorder2 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color5 = new Color() { Indexed = (UInt32Value)64U };

            topBorder2.Append(color5);
            //cambiado borderStyle = thick(gruesa) / thin(delgada)
            BottomBorder bottomBorder2 = new BottomBorder() { Style = BorderStyleValues.Thick };
            Color color6 = new Color() { Indexed = (UInt32Value)64U, Rgb = new HexBinaryValue("FF3333") };

            bottomBorder2.Append(color6);
            DiagonalBorder diagonalBorder2 = new DiagonalBorder();

            border2.Append(leftBorder2);
            border2.Append(rightBorder2);
            border2.Append(topBorder2);
            border2.Append(bottomBorder2);
            border2.Append(diagonalBorder2);

            borders1.Append(border1);
            borders1.Append(border2);

            #endregion

            #region StyleFormatCells
            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)1U };
            CellFormat cellFormat1 = new CellFormat() 
            { 
                NumberFormatId = (UInt32Value)0U, 
                FontId = (UInt32Value)0U, 
                FillId = (UInt32Value)0U, 
                BorderId = (UInt32Value)0U,
                FormatId = (UInt32Value)0U, // Agregado
            };

            cellStyleFormats1.Append(cellFormat1);

            CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)3U };
            //Maneja las celdas que no tienen datos
            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            //Maneja los datos del cuerpo
            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyBorder = true };
            //Maneja los encabezados
            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true };

            cellFormats1.Append(cellFormat2);
            cellFormats1.Append(cellFormat3);
            cellFormats1.Append(cellFormat4);

            CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)1U };
            CellStyle cellStyle1 = new CellStyle() { Name = "Normal", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

            cellStyles1.Append(cellStyle1);


            //No encuentro como modificar el fondo

            #endregion

            #region Diferencial Formato
            DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)0U };
            TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };

            StylesheetExtensionList stylesheetExtensionList1 = new StylesheetExtensionList();

            StylesheetExtension stylesheetExtension1 = new StylesheetExtension() { Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" };
            stylesheetExtension1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            X14.SlicerStyles slicerStyles1 = new X14.SlicerStyles() { DefaultSlicerStyle = "SlicerStyleLight1" };

            stylesheetExtension1.Append(slicerStyles1);

            StylesheetExtension stylesheetExtension2 = new StylesheetExtension() { Uri = "{9260A510-F301-46a8-8635-F512D64BE5F5}" };
            stylesheetExtension2.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            X15.TimelineStyles timelineStyles1 = new X15.TimelineStyles() { DefaultTimelineStyle = "TimeSlicerStyleLight1" };

            stylesheetExtension2.Append(timelineStyles1);

            stylesheetExtensionList1.Append(stylesheetExtension1);
            stylesheetExtensionList1.Append(stylesheetExtension2);
            #endregion

            #region StyleSheet
            stylesheet1.Append(fonts1);
            stylesheet1.Append(fills1);
            stylesheet1.Append(borders1);
            stylesheet1.Append(cellStyleFormats1);
            stylesheet1.Append(cellFormats1);
            stylesheet1.Append(cellStyles1);
            stylesheet1.Append(differentialFormats1);
            stylesheet1.Append(tableStyles1);
            stylesheet1.Append(stylesheetExtensionList1);
            #endregion

            workbookStylesPart1.Stylesheet = stylesheet1;
        }

        //recibe el listado por parametro, ¿agregar datos a Excel?, retorna un SheetData (hoja de datos)
        private SheetData GenerarHojadeDatos(TestModelList data)
        {
            SheetData sheetData1 = new SheetData();
            sheetData1.Append(CrearEncabezadosExcel());

            foreach (TestModel testmodel in data.testData)
            {
                Row partsRows = GeneraFilasDatos(testmodel);
                sheetData1.Append(partsRows);
            }
            return sheetData1;
        }

        //crear filas de encabezado en Excel.
        private Row CrearEncabezadosExcel()
        {
            Row workRow = new Row();
            //Cambiar nombre Encabezado y estilo
            //CreateCell("nombreEncabezado", 0U y 1U "textonormal" 2U "textonegrita")
            workRow.Append(CrearCeldas("Test Id", 2U));
            //Probando con el metodo GenerateRowForChildPartDetail
            workRow.Append(CrearCeldas("Test Id con Hola", 2U));
            workRow.Append(CrearCeldas("Test Name", 2U));
            workRow.Append(CrearCeldas("Test Description", 2U));
            workRow.Append(CrearCeldas("Test Date", 2U));
            //Si se olvida de colocar algun encabezado, la ultima columna no tendra encabezado
            return workRow;
        }

        //¿genera filas secundarias?. Devuelve los datos que se van a cargar en el Excel
        private Row GeneraFilasDatos(TestModel testmodel)
        {
            Row tRow = new Row();
            tRow.Append(CrearCeldas(testmodel.TestId.ToString()));
            //Probando: devuelve el id con un hola en negritra
            tRow.Append(CrearCeldas(testmodel.TestId.ToString() + "Hola"));
            tRow.Append(CrearCeldas(testmodel.TestName));
            tRow.Append(CrearCeldas(testmodel.TestDesc));
            tRow.Append(CrearCeldas(testmodel.TestDate.ToShortDateString()));

            return tRow;
        }

        //crear una celda pasando solo los datos de la celda y agrega un estilo predeterminado.
        private Cell CrearCeldas(string text)
        {
            Cell cell = new Cell();
            cell.StyleIndex = 1U;
            cell.DataType = ResolveCellDataTypeOnValue(text);
            cell.CellValue = new CellValue(text);
            return cell;
        }

        //crear una celda pasando datos de celda y estilo de celda.
        private Cell CrearCeldas(string text, uint styleIndex)
        {
            Cell cell = new Cell();
            cell.StyleIndex = styleIndex;
            cell.DataType = ResolveCellDataTypeOnValue(text);
            cell.CellValue = new CellValue(text);
            return cell;
        }

        //resolver el tipo de datos de valor numérico en una celda.
        private EnumValue<CellValues> ResolveCellDataTypeOnValue(string text)
        {
            int intVal;
            double doubleVal;
            if (int.TryParse(text, out intVal) || double.TryParse(text, out doubleVal))
            {
                return CellValues.Number;
            }
            else
            {
                return CellValues.String;
            }
        }


        /*PROBANDO COLORES*/
         
        private Stylesheet CreateStylesheet()
        {
            Stylesheet stylesheet1 = new Stylesheet();

            #region Font
            Fonts fonts = new Fonts() { Count = (UInt32Value)2U };

            // FontID = 0 (The Default Font)
            Font font0 = new Font();
            FontSize fontSize0 = new FontSize() { Val = 11D };
            Color color0 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName0 = new FontName() { Val = "Calibri" };

            font0.Append(fontSize0);
            font0.Append(color0);
            font0.Append(fontName0);

            // FontID = 1
            Font font1 = new Font();
            FontSize fontSize1 = new FontSize() { Val = 12D };
            FontName fontName1 = new FontName() { Val = "Calibri" };
            Color color1 = new Color() { Rgb = new HexBinaryValue("5B9BD5") };
            Bold bold1 = new Bold() { Val = true };

            // Foreground & Background not working
            ForegroundColor foregroundColor1 = new ForegroundColor() { Rgb = new HexBinaryValue("FF00FF00") };
            BackgroundColor backgroundColor1 = new BackgroundColor() { Rgb = new HexBinaryValue("FF00FF00") };


            font1.Append(fontSize1);
            font1.Append(fontName1);
            font1.Append(color1);
            font1.Append(bold1);
            font1.Append(foregroundColor1);
            font1.Append(backgroundColor1);

            fonts.Append(font0);
            fonts.Append(font1);

            int fontIndex = fonts.Count() - 1;
            #endregion

            #region CellBorders
            Borders borders = new Borders() { Count = (UInt32Value)2U };

            // Border ID = 0 (The Default Border)
            Border border0 = new Border();

            LeftBorder leftBorder0 = new LeftBorder();
            RightBorder rightBorder0 = new RightBorder();
            TopBorder topBorder0 = new TopBorder();
            BottomBorder bottomBorder0 = new BottomBorder();

            border0.Append(leftBorder0);
            border0.Append(rightBorder0);
            border0.Append(topBorder0);
            border0.Append(bottomBorder0);

            // Border ID = 1
            Border border1 = new Border();

            LeftBorder leftBorder = new LeftBorder() { Style = BorderStyleValues.Thick };
            RightBorder rightBorder = new RightBorder() { Style = BorderStyleValues.Thick };
            TopBorder topBorder = new TopBorder() { Style = BorderStyleValues.Thick };
            BottomBorder bottomBorder = new BottomBorder() { Style = BorderStyleValues.Thick };

            border1.Append(leftBorder);
            border1.Append(rightBorder);
            border1.Append(topBorder);
            border1.Append(bottomBorder);

            borders.Append(border0);
            borders.Append(border1);

            int borderIndex = borders.Count() - 1;
            #endregion

            #region Fill
            Fills fills = new Fills() { Count = (UInt32Value)3U };

            // Fill ID = 0 (The Default Fill)
            Fill fill0 = new Fill();
            PatternFill patternFill0 = new PatternFill() { PatternType = PatternValues.None };     // required, reserved by Excel
            fill0.Append(patternFill0);

            // Fill ID = 1 (The default fill of Gray)
            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.Gray125 };  // required, reserved by Excel
            fill1.Append(patternFill1);

            // Fill ID = 2
            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Solid };    // customized Pattern value
            ForegroundColor foregroundColorFill2 = new ForegroundColor() { Rgb = new HexBinaryValue("FF00FF00") };
            BackgroundColor backgroundColorFill2 = new BackgroundColor() { Rgb = new HexBinaryValue("5B9BD500") };

            patternFill2.Append(foregroundColorFill2);
            patternFill2.Append(backgroundColorFill2);
            fill2.Append(patternFill2);

            fills.Append(fill0);
            fills.Append(fill1);
            fills.Append(fill2);

            int fillIndex = fills.Count() - 1;
            #endregion

            #region Alignment
            Alignment alignment = new Alignment()
            {
                Horizontal = HorizontalAlignmentValues.Left,
                Vertical = VerticalAlignmentValues.Center
            };
            #endregion

            #region Cell Style Formats & Cell Formats
            CellFormats cellFormats0 = new CellFormats() { Count = (UInt32Value)1U };

            // Cell Format ID = 0 (The Default Cell Format)
            CellFormat cellFormat0 = new CellFormat()
            {
                NumberFormatId = (UInt32Value)0U,
                FontId = (UInt32Value)0U,
                FillId = (UInt32Value)0U,
                BorderId = (UInt32Value)0U,
                FormatId = (UInt32Value)0U
            };

            cellFormats0.Append(cellFormat0);

            // Cell Format ID = 1
            CellFormat cellFormat1 = new CellFormat()
            {
                NumberFormatId = (UInt32Value)1U,
                FontId = (UInt32Value)1U,
                FillId = (UInt32Value)2U,
                BorderId = (UInt32Value)1U,
                FormatId = (UInt32Value)1U,
                ApplyBorder = true,
                ApplyFont = true,
                ApplyFill = true,
            };
            cellFormats0.Append(cellFormat1);

            CellStyles cellStyles = new CellStyles() { Count = (UInt32Value)1U };
            CellStyle cellStyle = new CellStyle()
            {
                Name = "Normal",
                FormatId = (UInt32Value)0U,
                BuiltinId = (UInt32Value)1U
            };

            cellStyles.Append(cellStyle);
            #endregion

            #region Stylesheet
            stylesheet1.Append(fonts);
            stylesheet1.Append(fills);
            stylesheet1.Append(borders);
            stylesheet1.Append(cellFormats0);
            stylesheet1.Append(cellStyles);
            #endregion

            return stylesheet1;
        }

    }
}
