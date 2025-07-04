using AngleSharp.Common;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelXporter.Models;
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
        public static FileContentResult ExportToExcel<T>(this List<T> objList, string filename, StyleOptions? styleOptions = null)
        {
            Stream stream = new MemoryStream();
            using (var spreadsheetDocument = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = spreadsheetDocument.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                styleOptions ??= new StyleOptions();
                var styleSheet = CreateStylesheet(styleOptions);
                var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = styleSheet;
                stylesPart.Stylesheet.Save();

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
                        new Cell() { 
                            CellValue = new CellValue(name), 
                            DataType = CellValues.String,
                            StyleIndex = 1
                        }
                    );
                }

                worksheetPart.Worksheet.GetFirstChild<SheetData>().AppendChild(headerRow);
                foreach (var obj in objList)
                {
                    var values = obj.ToDictionary().Values;
                    var dataRow = new Row();
                    foreach (var value in values)
                    {
                        if (int.TryParse(value, out int result))
                        {
                            dataRow.Append(
                                new Cell() { 
                                    CellValue = new CellValue(result), 
                                    DataType = CellValues.Number,
                                    StyleIndex = 2
                                }
                            );
                        }
                        else
                        {
                            dataRow.Append(
                                new Cell() { 
                                    CellValue = new CellValue(value), 
                                    DataType = CellValues.String,
                                    StyleIndex = 2
                                }
                            );
                        }                        
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

        public static FileContentResult ExportToExcelMultipleSheets(this List<dynamic> sheetData, string filename, StyleOptions? styleOptions = null)
        {
            Stream stream = new MemoryStream();
            using (var spreadsheetDocument = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = spreadsheetDocument.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                styleOptions ??= new StyleOptions();
                var styleSheet = CreateStylesheet(styleOptions);
                var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = styleSheet;
                stylesPart.Stylesheet.Save();

                var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());

                int sheetId = 0;
                foreach (var objList in sheetData)
                {
                    sheetId++;
                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());

                    var sheet = new Sheet()
                    {
                        Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                        SheetId = (uint)sheetId,
                        Name = $"Sheet {sheetId}"
                    };
                    sheets.Append(sheet);
                    List<object> newObjList = new();
                    newObjList.AddRange(objList);
                    var columnNames = newObjList[0].ToDictionary().Keys;

                    var headerRow = new Row();
                    foreach (var name in columnNames)
                    {
                        headerRow.Append(
                        new Cell()
                        {
                            CellValue = new CellValue(name),
                            DataType = CellValues.String,
                            StyleIndex = 1
                        }
                    );
                    }

                    worksheetPart.Worksheet.GetFirstChild<SheetData>().AppendChild(headerRow);
                    foreach (var obj in newObjList)
                    {

                        var values = obj.ToDictionary().Values;
                        var dataRow = new Row();
                        foreach (var value in values)
                        {
                            if (int.TryParse(value, out int result))
                            {
                                dataRow.Append(
                                    new Cell()
                                    {
                                        CellValue = new CellValue(result),
                                        DataType = CellValues.Number,
                                        StyleIndex = 2
                                    }
                                );
                            }
                            else
                            {
                                dataRow.Append(
                                    new Cell()
                                    {
                                        CellValue = new CellValue(value),
                                        DataType = CellValues.String,
                                        StyleIndex = 2
                                    }
                                );
                            }
                        }
                        worksheetPart.Worksheet.GetFirstChild<SheetData>().AppendChild(dataRow);
                    }
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
                FileDownloadName = $"{fileDownloadName}.xlsx"
            };
        }

        private static Stylesheet CreateStylesheet(StyleOptions styleOptions)
        {
            var fonts = new Fonts(
                new Font(), // 0 - default
                new Font(   // 1 - header font
                    new Bold(),
                    new Color { Rgb = styleOptions.HeaderStyle.FontColorHex }),
                new Font(   // 2 - data font
                    new Color { Rgb = styleOptions.DefaultCellStyle.FontColorHex })
            );

            var fills = new Fills(
                new Fill(new PatternFill { PatternType = PatternValues.None }), // 0
                new Fill(new PatternFill { PatternType = PatternValues.Gray125 }), // 1
                new Fill(new PatternFill(new ForegroundColor { Rgb = styleOptions.HeaderStyle.BackgroundColorHex })
                {
                    PatternType = PatternValues.Solid
                }) // 2 - header fill
            );

            var borders = new Borders(new Border()); // default border
            uint borderId = 0;

            if (styleOptions.BorderStyle.ApplyBorders)
            {
                var border = new Border(
                    new LeftBorder { Style = styleOptions.BorderStyle.Style, Color = new Color { Rgb = styleOptions.BorderStyle.BorderColorHex } },
                    new RightBorder { Style = styleOptions.BorderStyle.Style, Color = new Color { Rgb = styleOptions.BorderStyle.BorderColorHex } },
                    new TopBorder { Style = styleOptions.BorderStyle.Style, Color = new Color { Rgb = styleOptions.BorderStyle.BorderColorHex } },
                    new BottomBorder { Style = styleOptions.BorderStyle.Style, Color = new Color { Rgb = styleOptions.BorderStyle.BorderColorHex } },
                    new DiagonalBorder()
                );
                borders.Append(border);
                borderId = 1;
            }

            var align = styleOptions.DefaultCellStyle.HorizontalAlignment switch
            {
                TextAlignment.Center => HorizontalAlignmentValues.Center,
                TextAlignment.Right => HorizontalAlignmentValues.Right,
                _ => HorizontalAlignmentValues.Left
            };

            var cellFormats = new CellFormats(
                new CellFormat(), // 0 - default
                new CellFormat // 1 - header
                {
                    FontId = 1,
                    FillId = 2,
                    BorderId = borderId,
                    ApplyFont = true,
                    ApplyFill = true,
                    ApplyBorder = styleOptions.BorderStyle.ApplyBorders
                },
                new CellFormat // 2 - data
                {
                    FontId = 2,
                    FillId = 0,
                    BorderId = borderId,
                    ApplyFont = true,
                    ApplyBorder = styleOptions.BorderStyle.ApplyBorders,
                    Alignment = new Alignment { Horizontal = align }
                }
            );

            return new Stylesheet(fonts, fills, borders, cellFormats);
        }



    }
}