using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;


namespace ExcelHelper
{
    public class ExcelHelper
    {
        public MemoryStream CreateExcel(List<DataTable> dts)
        {
            MemoryStream stream = new MemoryStream();
            SpreadsheetDocument Doc = SpreadsheetDocument.Create(stream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
            using (Doc)
            {
                Doc.PackageProperties.Creator = "PKUSCE";
                Doc.PackageProperties.Created = DateTime.Now;
                Doc.PackageProperties.LastModifiedBy = "PKUSCE";

                //创建WorkbookPart，在代码中主要使用这个相当于xml的root elements， spreadSheet.AddWorkbookPart()， 虽然是"Add"方法， 但你只能加一个。
                Doc.AddWorkbookPart();
                WorkbookStylesPart workbookStylesPart1 = Doc.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                workbookStylesPart1.Stylesheet = GenerateStyleSheet();
                Doc.WorkbookPart.Workbook = new Workbook();

                int sname = 0;
                foreach (DataTable dt in dts)
                {
                    List<string> ColumnNames = new List<string>();

                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        ColumnNames.Add(dt.Columns[i].ColumnName);
                    }

                    //添加工作表(Sheet)
                    WorksheetPart worksheetPart = InsertWorksheet(Doc.WorkbookPart, "测试" + sname.ToString());
                    sname++;

                    //创建列
                    //worksheetPart.Worksheet.InsertAfter(GenerateColumns(ColumnNames), worksheetPart.Worksheet.SheetProperties);
                    worksheetPart.Worksheet.InsertAfter(GenerateColumns(GetMax(dt)), worksheetPart.Worksheet.SheetProperties);

                    //创建多个工作表可共用的字符串容器
                    SharedStringTablePart shareStringPart = CreateSharedStringTablePart(Doc.WorkbookPart);

                    uint rowIndex = 1;
                    int ColumnIndex = 1;


                    //第一行、标题行
                    for (ColumnIndex = 1; ColumnIndex <= dt.Columns.Count; ColumnIndex++)
                    {
                        string name = dt.Columns[ColumnIndex - 1].ColumnName;
                        Cell cell = InsertCellInWorksheet(GetColumnName(ColumnIndex), rowIndex, worksheetPart);
                        //在共用字符串容器里插入一个字符串
                        int strIndex = InsertSharedStringItem(name, shareStringPart);
                        cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                        cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(strIndex.ToString());//注：这里要设置为目标字符串在SharedStringTablePart中的索引
                    }

                    rowIndex++;

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        ColumnIndex = 1;
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            String columnType = dt.Columns[j].DataType.ToString();
                            string _value = dt.Rows[i][j].ToString();
                            Cell cell = InsertCellInWorksheet(GetColumnName(ColumnIndex), rowIndex, worksheetPart);

                            //设置单元格的值
                            switch (columnType)
                            {
                                case "System.String":
                                    int strIndex = InsertSharedStringItem(_value, shareStringPart);
                                    cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                                    cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(strIndex.ToString());//注：这里要设置为目标字符串在SharedStringTablePart中的索引
                                    break;
                                case "System.Boolean"://布尔型                           
                                    break;
                                case "System.Int16":
                                case "System.Int64":
                                case "System.Byte":
                                case "System.Int32":
                                case "System.Decimal":
                                case "System.Double":
                                    cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                                    cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(_value);
                                    cell.StyleIndex = 7;
                                    break;
                                case "System.DateTime":
                                    cell.DataType = new EnumValue<CellValues>(CellValues.Date);
                                    cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue { Text = DateTime.Now.ToString("yyyy-MM-dd") };
                                    //cell.StyleIndex = 176;
                                    break;
                                case "System.DBNull"://空值处理                                
                                    break;
                                default:
                                    break;
                            }

                            ColumnIndex++;
                        }
                        rowIndex++;
                    }

                    worksheetPart.Worksheet.Save();
                }
            }


            return stream;
        }

        ///<summary>
        /// 获取列名称
        ///</summary>
        ///<param name="colIndex"></param>
        ///<returns></returns>
        private static string GetColumnName(int colIndex)
        {
            int dividend = colIndex;
            string columnName = String.Empty;
            int modifier;

            while (dividend > 0)
            {
                modifier = (dividend - 1) % 26;
                columnName =
                    Convert.ToChar(65 + modifier).ToString() + columnName;
                dividend = (int)((dividend - modifier) / 26);
            }
            return columnName;
        }


        /// <summary>
        /// 创建一个SharedStringTablePart(相当于各Sheet共用的存放字符串的容器)
        /// </summary>
        /// <param name="workbookPart"></param>
        /// <returns></returns>
        private static SharedStringTablePart CreateSharedStringTablePart(WorkbookPart workbookPart)
        {
            SharedStringTablePart shareStringPart = null;
            if (workbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                shareStringPart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
            }
            else
            {
                shareStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
            }
            return shareStringPart;
        }

        /// <summary>
        /// 插入worksheet
        /// </summary>
        /// <param name="workbookPart"></param>
        /// <returns></returns>
        private static WorksheetPart InsertWorksheet(WorkbookPart workbookPart, string sheetName = null)
        {
            //创建一个新的WorkssheetPart（后面将用它来容纳具体的Sheet）
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            newWorksheetPart.Worksheet.Save();

            //取得Sheet集合
            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            if (sheets == null)
            {
                sheets = workbookPart.Workbook.AppendChild<Sheets>(new Sheets());
            }

            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

            //得到Sheet的唯一序号
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            string sheetTempName = "Sheet" + sheetId;

            if (sheetName != null)
            {
                bool hasSameName = false;
                //检测是否有重名
                foreach (var item in sheets.Elements<Sheet>())
                {
                    if (item.Name == sheetName)
                    {
                        hasSameName = true;
                        break;
                    }
                }
                if (!hasSameName)
                {
                    sheetTempName = sheetName;
                }
            }

            //创建Sheet实例并将它与sheets关联
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetTempName };
            sheets.Append(sheet);
            workbookPart.Workbook.Save();

            return newWorksheetPart;
        }

        /// <summary>
        /// 插入共享字符串
        /// </summary>
        /// <param name="text"></param>
        /// <param name="shareStringPart"></param>
        /// <returns></returns>
        private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            //检测SharedStringTable是否存在，如果不存在，则创建一个
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
                shareStringPart.SharedStringTable.Count = 1;
                shareStringPart.SharedStringTable.UniqueCount = 1;
            }

            int i = 0;
            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }
                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }

        /// <summary>
        /// 向工作表插入一个单元格
        /// </summary>
        /// <param name="columnName">列名称</param>
        /// <param name="rowIndex">行索引</param>
        /// <param name="worksheetPart"></param>
        /// <returns></returns>
        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;//列的引用字符串，类似:"A3"或"B5"

            //如果指定的行存在，则直接返回该行，否则插入新行
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            //如果该行没有指定ColumnName的列，则插入新列，否则直接返回该列
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                //列必须按(字母)顺序插入，因此要先根据"列引用字符串"查找插入的位置
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }


        private Stylesheet GenerateStyleSheet()
        {
            return new Stylesheet(
                new Fonts(
                    new Font(                                                               // Index 0 – The default font.
                        new FontSize() { Val = 11 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Calibri" }),
                    new Font(                                                               // Index 1 – The bold font.
                        new Bold(),
                        new FontSize() { Val = 11 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Calibri" }),
                    new Font(                                                               // Index 2 – The Italic font.
                        new Italic(),
                        new FontSize() { Val = 11 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Calibri" }),
                    new Font(                                                               // Index 3 – The Times Roman font. with 16 size
                        new FontSize() { Val = 16 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Times New Roman" })
                ),
                new Fills(
                    new Fill(                                                           // Index 0 – The default fill.
                        new PatternFill() { PatternType = PatternValues.None }),
                    new Fill(                                                           // Index 1 – The default fill of gray 125 (required)
                        new PatternFill() { PatternType = PatternValues.Gray125 }),
                    new Fill(                                                           // Index 2 – The yellow fill.
                        new PatternFill(
                            new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "FFFFFF00" } }) { PatternType = PatternValues.Solid })
                ),
                new Borders(
                    new Border(                                                         // Index 0 – The default border.
                        new LeftBorder(),
                        new RightBorder(),
                        new TopBorder(),
                        new BottomBorder(),
                        new DiagonalBorder()),
                    new Border(                                                         // Index 1 – Applies a Left, Right, Top, Bottom border to a cell
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
                    new CellFormat() { FontId = 0, FillId = 0, BorderId = 0 },                          // Index 0 – The default cell style.  If a cell does not have a style index applied it will use this style combination instead
                    new CellFormat() { FontId = 1, FillId = 0, BorderId = 0, ApplyFont = true },       // Index 1 – Bold 
                    new CellFormat() { FontId = 2, FillId = 0, BorderId = 0, ApplyFont = true },       // Index 2 – Italic
                    new CellFormat() { FontId = 3, FillId = 0, BorderId = 0, ApplyFont = true },       // Index 3 – Times Roman
                    new CellFormat() { FontId = 0, FillId = 2, BorderId = 0, ApplyFill = true },       // Index 4 – Yellow Fill
                    new CellFormat(                                                                   // Index 5 – Alignment
                        new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }
                    ) { FontId = 0, FillId = 0, BorderId = 0, ApplyAlignment = true },
                    new CellFormat() { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true },      // Index 6 – Border
                    new CellFormat() { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true, NumberFormatId = 10 } // Index 7 – Border
                )
            ); // return
        }




        //        public class ExcelNumberFormat
        //    {
        //        private static Dictionary<uint, string> _builtInFormats_Global = new Dictionary<uint, string>()
        //        {
        //            {0, "General"},
        //            {1, "0"},
        //            {2, "0.00"},
        //            {3, "#,##0"},
        //            {4, "#,##0.00"},
        //            {9, "0%"},
        //            {10, "0.00%"},
        //            {11, "0.00E+00"},
        //            {12, "# ?/?"},
        //            {13, "# ??/??"},
        //            {14, "m/d/yyyy"},
        //            {15, "d-mmm-yy"},
        //            {16, "d-mmm"},
        //            {17, "mmm-yy"},
        //            {18, "h:mm AM/PM"},
        //            {19, "h:mm:ss AM/PM"},
        //            {20, "h:mm"},
        //            {21, "h:mm:ss"},
        //            {22, "m/d/yy h:mm"},
        //            {37, "#,##0 ;(#,##0)"},
        //            {38, "#,##0 ;[Red](#,##0)"},
        //            {39, "#,##0.00;(#,##0.00)"},
        //            {40, "#,##0.00;[Red](#,##0.00)"},
        //            {45, "mm:ss"},
        //            {46, "[h]:mm:ss"},
        //            {47, "mmss.0"},
        //            {48, "##0.0E+0"},
        //            {49, "@"}
        //        };

        //        private DocumentStyles _styles;
        //        private IStylable _stylable;

        //        internal uint NumFmtId { get; set; }

        //        internal ExcelNumberFormat(IStylable stylable, DocumentStyles styles, uint numFmtId)
        //        {
        //            _stylable = stylable;
        //            _styles = styles;
        //            NumFmtId = numFmtId;
        //        }

        //        public string Format
        //        {
        //            get
        //            {
        //                if (_builtInFormats_Global.ContainsKey(NumFmtId))
        //                    return _builtInFormats_Global[NumFmtId];
        //                NumberingFormat numFmt = _styles.GetNumberingFormat(NumFmtId);
        //                return numFmt.FormatCode;
        //            }
        //            set
        //            {
        //                uint newNumFmtId;
        //                KeyValuePair<uint, string> builtInFmt = (from i in _builtInFormats_Global
        //                                                         where i.Value == value
        //                                                         select i).FirstOrDefault();
        //                if (builtInFmt.Value == value)
        //                {
        //                    newNumFmtId = builtInFmt.Key;
        //                }
        //                else
        //                {
        //                    NumberingFormat numFmt = new NumberingFormat() { FormatCode = value };
        //                    newNumFmtId = _styles.EnsureCustomNumberingFormat(numFmt);
        //                }
        //                if (newNumFmtId != NumFmtId)
        //                {
        //                    NumFmtId = newNumFmtId;
        //                    if (_stylable != null)
        //                        _stylable.Style.NumberFormat = this;
        //                }
        //            }
        //        }
        //    }
        //}
        private static void GenerateWorkbookStylesPart1Content(WorkbookStylesPart workbookStylesPart)
        {

            /*
             *  在显示cell是通过StyleIndex 来关联 cellXfs的Index 来改变cell 的显示样式， 
                注意， 这个index只能从1 开始，因此需要在cellXfs中加两个CellFormat子节点， 
                我们这里要设置 wrap text， 因此在第二个节点设置applyAlignment 并设wrap Text ="1". 
                怎么设置cell的 font，答案就是加一个font 子节点到fonts，
                得到index， 再加一个cellformat 子节点 并设置fontid 为刚加的font的index。 把这个cellformat的id 给 要设置的cell的StyleIndex。
             */


            Stylesheet stylesheet = new Stylesheet();

            //在创建stylesheet时， 必须创建fonts， Fills，Borders 和cellXfs（CellFormats） 四个节点
            Fonts fonts = new Fonts() { Count = (UInt32Value)2U, KnownFonts = false };
            Fills fills = new Fills() { Count = (UInt32Value)2U };
            Borders borders = new Borders() { Count = (UInt32Value)1U };
            CellFormats cellFormats = new CellFormats() { Count = (UInt32Value)2U };
            CellStyleFormats cellStyleFormats = new CellStyleFormats() { Count = (UInt32Value)1U };
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };


            Font font1 = new Font();
            FontSize fontSize1 = new FontSize() { Val = 11D };
            Color color1 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName1 = new FontName() { Val = "宋体" };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet1 = new FontCharSet() { Val = 134 };
            DocumentFormat.OpenXml.Spreadsheet.FontScheme fontScheme1 = new DocumentFormat.OpenXml.Spreadsheet.FontScheme() { Val = FontSchemeValues.Minor };

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontCharSet1);
            font1.Append(fontScheme1);

            Font font2 = new Font();
            FontSize fontSize2 = new FontSize() { Val = 9D };
            FontName fontName2 = new FontName() { Val = "宋体" };
            FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet2 = new FontCharSet() { Val = 134 };
            DocumentFormat.OpenXml.Spreadsheet.FontScheme fontScheme2 = new DocumentFormat.OpenXml.Spreadsheet.FontScheme() { Val = FontSchemeValues.Minor };

            font2.Append(fontSize2);
            font2.Append(fontName2);
            font2.Append(fontFamilyNumbering2);
            font2.Append(fontCharSet2);
            font2.Append(fontScheme2);

            fonts.Append(font1);
            fonts.Append(font2);



            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };
            fill1.Append(patternFill1);
            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };
            fill2.Append(patternFill2);
            fills.Append(fill1);
            fills.Append(fill2);



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
            borders.Append(border1);




            Alignment alignment1 = new Alignment() { Vertical = VerticalAlignmentValues.Center };
            cellFormat1.Append(alignment1);

            cellStyleFormats.Append(cellFormat1);


            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };

            Alignment alignment2 = new Alignment() { Vertical = VerticalAlignmentValues.Center };

            cellFormat2.Append(alignment2);

            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)14U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true };
            Alignment alignment3 = new Alignment() { Vertical = VerticalAlignmentValues.Center };


            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = 176, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };
            cellFormat3.Append(alignment3);

            cellFormats.Append(cellFormat2);
            cellFormats.Append(cellFormat3);
            cellFormats.Append(cellFormat4);

            CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)1U };
            CellStyle cellStyle1 = new CellStyle() { Name = "常规", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

            cellStyles1.Append(cellStyle1);
            DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)0U };
            TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };

            stylesheet.Append(fonts);
            stylesheet.Append(fills);
            stylesheet.Append(borders);
            stylesheet.Append(cellStyleFormats);
            stylesheet.Append(cellFormats);
            stylesheet.Append(cellStyles1);
            stylesheet.Append(differentialFormats1);
            stylesheet.Append(tableStyles1);
            //stylesheet.Append(stylesheetExtensionList1);

            workbookStylesPart.Stylesheet = stylesheet;
        }


        public static Columns GenerateColumns(List<int> ColumnNames)
        {
            double MaxWidth = 11;
            Columns columns = new Columns();
            for (int i = 0; i < ColumnNames.Count; i++)
            {
                UInt32Value index = new UInt32Value((uint)i + 1);               
                var width = Math.Truncate((ColumnNames[i] * MaxWidth + 5.0) / MaxWidth * 256) / 256 +2;
                Column col = new Column() { Min = index, Max = index, Width = width, CustomWidth = true, BestFit = true };
                columns.Append(col);
            }
            return columns;
        }



        public static Columns GenerateColumns(List<string> ColumnNames)
        {
            double MaxWidth = 11;
            Columns columns = new Columns();
            for (int i = 0; i < ColumnNames.Count; i++)
            {
                UInt32Value index = new UInt32Value((uint)i + 1);
                string colName = ColumnNames[i];
                int len = GetTrueLength(colName);
                var width = Math.Truncate((len * MaxWidth + 5.0) / MaxWidth * 256) / 256;
                Column col = new Column() { Min = index, Max = index, Width = width, CustomWidth = true, BestFit = true };
                columns.Append(col);
            }
            return columns;
        }


        /// <summary>
        /// 获取字符串的真实长度。解决中文和英文的之间不同
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        private static int GetTrueLength(string input)
        {
            // str 字符串
            // return 字符串的字节长度
            int lenTotal = 0;
            int n = input.Length;
            string strWord = "";
            int asc;
            for (int i = 0; i < n; i++)
            {
                strWord = input.Substring(i, 1);
                asc = Convert.ToChar(strWord);
                if (asc < 0 || asc > 127)
                    lenTotal = lenTotal + 2;
                else
                    lenTotal = lenTotal + 1;
            }
            return lenTotal;
        }


        /// <summary>
        /// 获取最长的那一列
        /// </summary>
        private static List<int> GetMax(DataTable dt)
        {
            List<int> t = new List<int>();
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                DataColumn dc = dt.Columns[i];

                DataColumn maxLengthColumn = new DataColumn();
                maxLengthColumn.ColumnName = "MaxLength";
                maxLengthColumn.Expression = "len(convert('" + dc.ColumnName + "','System.String'))";
                dt.Columns.Add(maxLengthColumn);
                object maxLength = dt.Compute("max(MaxLength)", "true");
                dt.Columns.Remove(maxLengthColumn);
                int len = Convert.ToInt32(maxLength);
                if (len > GetTrueLength(dc.ColumnName))
                {
                    t.Add(len);
                }
                else
                {
                    t.Add(GetTrueLength(dc.ColumnName));
                }
            }

            return t;
        }
    }
}
