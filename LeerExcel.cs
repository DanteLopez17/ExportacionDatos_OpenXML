using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExportacionDatos_OpenXML
{
    public class LeerExcel
    {
        // The SAX approach.
        public List<string> LeerExcelSAX(string fileName)
        {
            List<string> datos = new List<string>();
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

                OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
                string text;
                while (reader.Read())
                {
                    if (reader.ElementType == typeof(CellValue))
                    {
                        text = reader.GetText();
                        Console.Write(text + "-");
                        datos.Add(text);
                    }
                }
                Console.WriteLine();
                Console.ReadKey();
            }

            return datos;
        }
    }
}
