using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelHelper
{
    public class ExcelHelper
    {
        public static byte[] GenerteExcelFromDataTable(DataTable dt)
        {
            byte[] bFile = null;
            MemoryStream mem = new MemoryStream();
            SpreadsheetDocument ExcelDoc = SpreadsheetDocument.Create(mem, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
            ExcelDoc.PackageProperties.Creator = "PKUSCE";
            ExcelDoc.PackageProperties.Created = DateTime.Now;
            ExcelDoc.PackageProperties.LastModifiedBy = "PKUSCE";

            string sheetName = dt.TableName;
            using (ExcelDoc)
            {
                ExcelDoc.AddWorkbookPart();
                ExcelDoc.WorkbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart1 = ExcelDoc.WorkbookPart.AddNewPart<WorksheetPart>();
                worksheetPart1.Worksheet = new Worksheet();
                SheetData sheetData = new SheetData();
                FillData(sheetData, dt);
                worksheetPart1.Worksheet.AppendChild(AutoFit(sheetData));
                worksheetPart1.Worksheet.AppendChild(sheetData);
                WorkbookStylesPart workbookStylesPart = ExcelDoc.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                FillStyle(workbookStylesPart);

                ExcelDoc.WorkbookPart.WorksheetParts.ElementAt(0).Worksheet.Save();

                ExcelDoc.WorkbookPart.Workbook.AppendChild(new Sheets());
                ExcelDoc.WorkbookPart.Workbook.GetFirstChild<Sheets>().AppendChild(new Sheet()
                {
                    Id = ExcelDoc.WorkbookPart.GetIdOfPart(ExcelDoc.WorkbookPart.WorksheetParts.First()),
                    SheetId = 1,
                    Name = sheetName
                });

                ExcelDoc.WorkbookPart.Workbook.Save();
            }
            bFile = mem.ToArray();
            return bFile;
        }

        /// <summary>
        /// 填充数据
        /// </summary>
        /// <param name="sheetData"></param>
        /// <param name="dt"></param>
        private static void FillData(SheetData sheetData, DataTable dt)
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
            #region
            //数据行
            for (rowIndex = 0; rowIndex < dt.Rows.Count; rowIndex++)
            {
                row = new Row() { RowIndex = (uint)(rowIndex + 2) };//从第二行开始
                for (int columnIndex = 0; columnIndex < dt.Columns.Count; columnIndex++)
                {
                    Cell cell = new Cell();

                    String columnType = dt.Columns[columnIndex].DataType.ToString();
                    string _value = dt.Rows[rowIndex][columnIndex].ToString();

                    switch (columnType)
                    {
                        case "System.String":
                            cell.DataType = new EnumValue<CellValues>(CellValues.String);
                            cell.CellValue = new CellValue(_value);
                            cell.StyleIndex = 6;
                            break;
                        case "System.Boolean"://布尔型                           
                            break;
                        case "System.Int16":
                        case "System.Int64":
                        case "System.Byte":
                        case "System.Int32":
                            cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            cell.CellValue = new CellValue(_value);
                            cell.StyleIndex = 6;
                            break;
                        case "System.Decimal":
                        case "System.Double":
                            cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            cell.CellValue = new CellValue(_value);
                            cell.StyleIndex = 8;
                            break;
                        case "System.DateTime":
                            cell.DataType = new EnumValue<CellValues>(CellValues.String);
                            //cell.CellValue = new CellValue( DateTime.Parse(_value).ToOADate().ToString(CultureInfo.CurrentCulture));
                            cell.CellValue = new CellValue(DateTime.Parse(_value).ToShortDateString());
                            cell.StyleIndex = 6;
                            break;
                        case "System.DBNull"://空值处理                                
                            break;
                        default:
                            break;

                    }                
                    row.Append(cell);
                }             
                sheetData.Append(row);              
            }
        #endregion
        }

        /// <summary>
        /// 设置单元格列宽
        /// </summary>
        /// <param name="sheetData"></param>
        /// <returns></returns>
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

                Column col = new Column() { BestFit = true, Min = (UInt32)(item.Key + 1), Max = (UInt32)(item.Key + 1), CustomWidth = true, Width = (DoubleValue)width + 2 };
                columns.Append(col);
            }
            return columns;
        }

        /// <summary>
        /// 计算每列宽度
        /// </summary>
        /// <param name="sheetData"></param>
        /// <returns></returns>
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
        private static void FillStyle(WorkbookStylesPart workbookStylesPart)
        {

            #region FormatCode

            //ID  FORMAT CODE
            //0   General
            //1   0
            //2   0.00
            //3   #,##0
            //4   #,##0.00
            //9   0%
            //10  0.00%
            //11  0.00E+00
            //12  # ?/?
            //13  # ??/??
            //14  d/m/yyyy
            //15  d-mmm-yy
            //16  d-mmm
            //17  mmm-yy
            //18  h:mm tt
            //19  h:mm:ss tt
            //20  H:mm
            //21  H:mm:ss
            //22  m/d/yyyy H:mm
            //37  #,##0 ;(#,##0)
            //38  #,##0 ;[Red](#,##0)
            //39  #,##0.00;(#,##0.00)
            //40  #,##0.00;[Red](#,##0.00)
            //45  mm:ss
            //46  [h]:mm:ss
            //47  mmss.0
            //48  ##0.0E+0
            //49  @
            #endregion

            Fonts fonts = new Fonts();
            #region 定义字体

            //Index=0 默认字体
            Font DefaultFont = new Font(
                new FontSize() { Val = 11 },
                new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                new FontName() { Val = "Calibri" });

            //Index=1 粗体字 
            Font BoldFont = new Font(
                new Bold(),
                new FontSize() { Val = 11 },
                new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                new FontName() { Val = "Calibri" });

            // Index=2 斜体字
            Font ItalicFont = new Font(
                new Italic(),
                new FontSize() { Val = 11 },
                new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                new FontName() { Val = "Calibri" });

            //Index=3 罗马字体
            Font TimesNewRomanFont = new Font(
               new Italic(),
               new FontSize() { Val = 11 },
               new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
               new FontName() { Val = "Times New Roman" });

            //Index=4 小字体
            Font SmallFont = new Font(
                new FontSize() { Val = 9 },
                new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                new FontName() { Val = "Calibri" });

            fonts.Append(DefaultFont);
            fonts.Append(BoldFont);
            fonts.Append(ItalicFont);
            fonts.Append(TimesNewRomanFont);
            fonts.Append(SmallFont);
            #endregion

            Fills fills = new Fills();
            #region 定义填充
            //默认 Index 0
            Fill Defaultfill = new Fill(new PatternFill() { PatternType = PatternValues.None });

            //DarkGray填充  Index1
            Fill DarkGrayfill = new Fill(new PatternFill() { PatternType = PatternValues.DarkGray });

            //前景色填充 色值#4876FF Index3
            Fill ForegroundColor = new Fill(
                new PatternFill(new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "4876FF" } }
                    ) { PatternType = PatternValues.Solid });

            fills.Append(Defaultfill);
            fills.Append(DarkGrayfill);
            fills.Append(ForegroundColor);
            #endregion

            Borders borders = new Borders();
            #region 定义边框
            //默认边框 Index 0 - 
            Border DefaultBorder = new Border(
                        new LeftBorder(),
                        new RightBorder(),
                        new TopBorder(),
                        new BottomBorder(),
                        new DiagonalBorder());

            //上下左右边框 Index 1
            Border CustomBorder = new Border(
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
                                    new DiagonalBorder());
            borders.Append(DefaultBorder);
            borders.Append(CustomBorder);
            #endregion

            NumberingFormats nfs = new NumberingFormats();
            #region 定义格式
            uint iExcelIndex = 164;
            //日期格式
            NumberingFormat nfDate = new NumberingFormat
            {
                FormatCode = StringValue.FromString(@"yyyy/m/d;@"),
                NumberFormatId = iExcelIndex++
            };
            //数字格式，保留一位小数
            NumberingFormat nfDouble = new NumberingFormat
            {
                FormatCode = StringValue.FromString("0.0"),
                NumberFormatId = iExcelIndex++
            };

            nfs.Append(nfDate);
            nfs.Append(nfDouble);
            #endregion

            CellFormats cfs = new CellFormats();
            #region 定义单元格样式
            //各种自由组合。这里只组合所需要的样式


            //默认样式 Index 0 
            CellFormat DefaultCellFormat = new CellFormat() { FontId = 0, FillId = 0, BorderId = 1 };

            //Index 1  粗体字 字号11
            CellFormat CellFormat1 = new CellFormat() { FontId = 1, FillId = 0, BorderId = 1, ApplyFont = true };

            //Index2   斜体字 字号11
            CellFormat CellFormat2 = new CellFormat() { FontId = 2, FillId = 0, BorderId = 1, ApplyFont = true };

            //Index3  罗马字体 字号11
            CellFormat CellFormat3 = new CellFormat() { FontId = 3, FillId = 0, BorderId = 1, ApplyFont = true };

            //Index4  背景填充 字号11
            CellFormat CellFormat4 = new CellFormat() { FontId = 0, FillId = 1, BorderId = 1, ApplyFill = true };

            //Index5 应用于Excel的标题行  居中对齐、边框，蓝色背景 字号11
            CellFormat CellFormat5 = new CellFormat(
                new Alignment()
                {
                    Horizontal = HorizontalAlignmentValues.Center,
                    Vertical = VerticalAlignmentValues.Center
                })
                {
                    FontId = 1,
                    FillId = 2,
                    BorderId = 1,
                    ApplyAlignment = true
                };

            //Index6 应用于Excel的标题数据行,文本类型 居中对齐、填充，边框 字号11
            CellFormat CellFormat6 = new CellFormat(
                new Alignment()
                {
                    Horizontal = HorizontalAlignmentValues.Center,
                    Vertical = VerticalAlignmentValues.Center
                })
            {
                FontId = 4,
                FillId = 0,
                BorderId = 1,
                ApplyAlignment = true
            };

            //index7 Excel数据行，日期类型
            CellFormat CellFormat7 = new CellFormat(
                new Alignment()
                {
                    Horizontal = HorizontalAlignmentValues.Center,
                    Vertical = VerticalAlignmentValues.Center
                })
            {
                FontId = 4,
                FillId = 0,
                BorderId = 1,
                NumberFormatId=nfDate.NumberFormatId,
                ApplyNumberFormat = BooleanValue.FromBoolean(true),
                ApplyAlignment = BooleanValue.FromBoolean(true)
            };

            //index8 Excel数据行，浮点数，保留1位小数
            CellFormat CellFormat8 = new CellFormat(
               new Alignment()
               {
                   Horizontal = HorizontalAlignmentValues.Center,
                   Vertical = VerticalAlignmentValues.Center
               })
            {
                FontId = 4,
                FillId = 0,
                BorderId = 1,
                NumberFormatId =nfDouble.NumberFormatId,
                ApplyNumberFormat = BooleanValue.FromBoolean(true),
                ApplyAlignment = BooleanValue.FromBoolean(true)
            };

            

            cfs.Append(DefaultCellFormat);
            cfs.Append(CellFormat1);
            cfs.Append(CellFormat2);
            cfs.Append(CellFormat3);
            cfs.Append(CellFormat4);
            cfs.Append(CellFormat5);
            cfs.Append(CellFormat6);
            cfs.Append(CellFormat7);
            cfs.Append(CellFormat8);
            #endregion

            Stylesheet stylesheet = new Stylesheet()
            {
                Fonts = fonts,
                Fills = fills,                        
                Borders = borders,
                NumberingFormats=nfs,
                CellFormats = cfs                
            };

            workbookStylesPart.Stylesheet = stylesheet;        
        }
    }
}