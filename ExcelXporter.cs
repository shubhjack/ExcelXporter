using AngleSharp.Common;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Mvc;

namespace ExcelXporter
{
    public static class ExcelXporter
    {
        /// <summary>
        /// Export any data model list to an excel file
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="objList"></param>
        /// <param name="filename"></param>
        /// <returns>.xlsx file</returns>
        public static FileContentResult ExportToExcel<T>(List<T> objList, string filename)
        {
            Stream stream = new MemoryStream();
            using (var spreadsheetDocument = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = spreadsheetDocument.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());

                var sheet = new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Sheet 1"
                };
                sheets.Append(sheet);

                var columnNames = objList[0].ToDictionary().Keys;

                var headerRow = new Row();
                foreach (var name in columnNames)
                {
                    headerRow.Append(
                        new Cell() { CellValue = new CellValue(name), DataType = CellValues.String }
                    );
                }

                worksheetPart.Worksheet.GetFirstChild<SheetData>().AppendChild(headerRow);
                foreach (var obj in objList)
                {
                    var values = obj.ToDictionary().Values;
                    var dataRow = new Row();
                    foreach (var value in values)
                    {
                        dataRow.Append(
                            new Cell() { CellValue = new CellValue(value), DataType = CellValues.String }
                        );
                    }
                    worksheetPart.Worksheet.GetFirstChild<SheetData>().AppendChild(dataRow);
                }
                workbookPart.Workbook.Save();
            }
            stream.Position = 0;
            byte[] bytes;
            using (var memoryStream = new MemoryStream())
            {
                stream.CopyTo(memoryStream);
                bytes = memoryStream.ToArray();
            }

            var fileDownloadName = filename.Replace(".xls", string.Empty).Replace(".xlsx", string.Empty).Trim();
            return new FileContentResult(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            {
                FileDownloadName =  $"{fileDownloadName}.xlsx"
            };
        }
    }
}