using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
 
 namespace OpenXmlGenerateExcelTest
 {
     public class test
     {         
         public static void CreateSpreadSheet(DataTable dt)
         {
             string fileName = @"C:\Users\LCQ\Desktop\test.xlsx";
             string sheetName = dt.TableName;
             using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
             {
                 spreadSheet.AddWorkbookPart();
                 spreadSheet.WorkbookPart.Workbook = new Workbook();
 
                 WorksheetPart worksheetPart1 = spreadSheet.WorkbookPart.AddNewPart<WorksheetPart>();
                 worksheetPart1.Worksheet = new Worksheet();
                 SheetData sheetData = new SheetData();
                 ProductData(sheetData,dt);
                 worksheetPart1.Worksheet.AppendChild(AutoFit(sheetData));
                 worksheetPart1.Worksheet.AppendChild(sheetData);
                 WorkbookStylesPart workbookStylesPart = spreadSheet.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                 GenerateWorkbookStylesPartContent(workbookStylesPart);
 
                 spreadSheet.WorkbookPart.WorksheetParts.ElementAt(0).Worksheet.Save();
 
                 spreadSheet.WorkbookPart.Workbook.AppendChild(new Sheets());
                 spreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>().AppendChild(new Sheet()
                 {
                     Id = spreadSheet.WorkbookPart.GetIdOfPart(spreadSheet.WorkbookPart.WorksheetParts.First()),
                     SheetId = 1,
                     Name = sheetName
                 });
 
                 spreadSheet.WorkbookPart.Workbook.Save();
             }
         }

         private static void ProductData(SheetData sheetData, DataTable dt)
         {
             //标题行
             int rowIndex = 1;
             Row row = new Row() { RowIndex = (uint)rowIndex };
             for (int i = 0; i < dt.Columns.Count; i++)
             {
                 Cell cell = new Cell();
                 cell.CellValue = new CellValue(dt.Columns[i].ColumnName.ToString());
                 cell.DataType = new EnumValue<CellValues>(CellValues.String);
                 cell.StyleIndex = 5;
                 row.Append(cell);
             }
             sheetData.Append(row);

             //数据行
             for (rowIndex = 0; rowIndex < dt.Rows.Count; rowIndex++)
             {
                 row = new Row() { RowIndex =(uint) (rowIndex+2) };//从第二行开始
                 for (int columnIndex = 0; columnIndex < dt.Columns.Count; columnIndex++)
                 {
                     Cell cell = new Cell();
                     cell.CellValue = new CellValue(dt.Rows[rowIndex][columnIndex].ToString());
                     cell.DataType = new EnumValue<CellValues>(CellValues.String);
                     cell.StyleIndex = 5;
                     row.Append(cell);
                 }
                 sheetData.Append(row);
             }            
         }
 
         private static Columns AutoFit(SheetData sheetData)
         {
             var maxColWidth = GetMaxCharacterWidth(sheetData);
 
             Columns columns = new Columns();
 
             double maxWidth = 7;
             foreach (var item in maxColWidth)
             {
                 /*三种单位宽度公式*/
                 double width = Math.Truncate((item.Value * maxWidth + 5) / maxWidth * 256) / 256;
                 double pixels = Math.Truncate(((256 * width + Math.Truncate(128 / maxWidth)) / 256) * maxWidth);
                 double charWidth = Math.Truncate((pixels - 5) / maxWidth * 100 + 0.5) / 100;

                 Column col = new Column() { BestFit = true, Min = (UInt32)(item.Key + 1), Max = (UInt32)(item.Key + 1), CustomWidth = true, Width = (DoubleValue)width+2 };
                 columns.Append(col);
             }
             return columns;
         }
 
         private static Dictionary<int, int> GetMaxCharacterWidth(SheetData sheetData)
         {
             Dictionary<int, int> maxColWidth = new Dictionary<int, int>();
             var rows = sheetData.Elements<Row>();
             foreach (var r in rows)
             {
                 var cells = r.Elements<Cell>().ToArray();
                 for (int i = 0; i < cells.Length; i++)
                 {
                     var cell = cells[i];
                     var cellValue = cell.CellValue == null ? string.Empty : cell.CellValue.InnerText;
                     var cellTextLength = cellValue.Length;
                     if (maxColWidth.ContainsKey(i))
                     {
                         var current = maxColWidth[i];
                         if (cellTextLength > current)
                         {
                             maxColWidth[i] = cellTextLength;
                         }
                     }
                     else
                     {
                         maxColWidth.Add(i, cellTextLength);
                     }
                 }
             }
             return maxColWidth;
         }


         /// <summary>
         /// 创建样式
         /// </summary>
         /// <param name="workbookStylesPart"></param>
         private static void GenerateWorkbookStylesPartContent(WorkbookStylesPart workbookStylesPart)
         {

             /*
              *  在显示cell是通过StyleIndex 来关联 cellXfs的Index 来改变cell 的显示样式， 
                 注意， 这个index只能从1 开始，因此需要在cellXfs中加两个CellFormat子节点， 
                 我们这里要设置 wrap text， 因此在第二个节点设置applyAlignment 并设wrap Text ="1". 
                 怎么设置cell的 font，答案就是加一个font 子节点到fonts，
                 得到index， 再加一个cellformat 子节点 并设置fontid 为刚加的font的index。 把这个cellformat的id 给 要设置的cell的StyleIndex。
              */


             Stylesheet stylesheet = new Stylesheet(
                 new Fonts(
                     new Font(                                                               // Index 0 - The default font.
                         new FontSize() { Val = 11 },
                         new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                         new FontName() { Val = "Calibri" }),
                     new Font(                                                               // Index 1 - The bold font.
                         new Bold(),
                         new FontSize() { Val = 11 },
                         new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                         new FontName() { Val = "Calibri" }),
                     new Font(                                                               // Index 2 - The Italic font.
                         new Italic(),
                         new FontSize() { Val = 11 },
                         new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                         new FontName() { Val = "Calibri" }),
                     new Font(                                                               // Index 2 - The Times Roman font. with 16 size
                         new FontSize() { Val = 16 },
                         new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                         new FontName() { Val = "Times New Roman" })
                 ),
                 new Fills(
                     new Fill(                                                           // Index 0 - The default fill.
                         new PatternFill() { PatternType = PatternValues.None }),
                     new Fill(                                                           // Index 1 - The default fill of gray 125 (required)
                         new PatternFill() { PatternType = PatternValues.Gray125 }),
                     new Fill(                                                           // Index 2 - The yellow fill.
                         new PatternFill(
                             new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "FFFFFF00" } }
                         ) { PatternType = PatternValues.Solid })
                 ),
                 new Borders(
                     new Border(                                                         // Index 0 - The default border.
                         new LeftBorder(),
                         new RightBorder(),
                         new TopBorder(),
                         new BottomBorder(),
                         new DiagonalBorder()),
                     new Border(                                                         // Index 1 - Applies a Left, Right, Top, Bottom border to a cell
                         new LeftBorder(
                             new Color() { Auto = true }
                         ) { Style = BorderStyleValues.Thin },
                         new RightBorder(
                             new Color() { Auto = true }
                         ) { Style = BorderStyleValues.Thin },
                         new TopBorder(
                             new Color() { Auto = true }
                         ) { Style = BorderStyleValues.Thin },
                         new BottomBorder(
                             new Color() { Auto = true }
                         ) { Style = BorderStyleValues.Thin },
                         new DiagonalBorder())
                 ),
                 new CellFormats(
                     new CellFormat() { FontId = 0, FillId = 0, BorderId = 0 },                          // Index 0 - The default cell style.  If a cell does not have a style index applied it will use this style combination instead
                     new CellFormat() { FontId = 1, FillId = 0, BorderId = 0, ApplyFont = true },       // Index 1 - Bold 
                     new CellFormat() { FontId = 2, FillId = 0, BorderId = 0, ApplyFont = true },       // Index 2 - Italic
                     new CellFormat() { FontId = 3, FillId = 0, BorderId = 0, ApplyFont = true },       // Index 3 - Times Roman
                     new CellFormat() { FontId = 0, FillId = 2, BorderId = 0, ApplyFill = true },       // Index 4 - Yellow Fill
                     new CellFormat(                                                                   // Index 5 - Alignment
                         new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }
                     ) { FontId = 0, FillId = 0, BorderId = 0, ApplyAlignment = true },
                     new CellFormat() { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true }      // Index 6 - Border
                 )
             );

             workbookStylesPart.Stylesheet = stylesheet;
         }
     }
 }