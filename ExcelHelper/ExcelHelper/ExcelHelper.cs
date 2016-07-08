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

                //创建样式部分
                WorkbookStylesPart workbookStylesPart1 = Doc.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                workbookStylesPart1.Stylesheet = GenerateStyleSheet();

                Doc.WorkbookPart.Workbook = new Workbook();
                
                //创建多个工作表可共用的字符串容器
                SharedStringTablePart shareStringPart = CreateSharedStringTablePart(Doc.WorkbookPart);

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
                    List<int> ColumnWidths = GetMax(dt);
                    worksheetPart.Worksheet.InsertAfter(GenerateColumns(ColumnWidths), worksheetPart.Worksheet.SheetProperties);



                    uint rowIndex = 1;
                    int ColumnIndex = 1;

                    //第一行、标题行
                    StringValue firstCellReference = null;
                    StringValue LastCellReference = null;
                    Cell firstCell = InsertCellInWorksheet(GetColumnName(1), rowIndex, worksheetPart);
                    Cell lastCell = InsertCellInWorksheet(GetColumnName(dt.Columns.Count), rowIndex, worksheetPart);
                    int strIndex0 = InsertSharedStringItem("合并单元格", shareStringPart);
                    firstCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                    firstCell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(strIndex0.ToString());
                    firstCell.StyleIndex = 7;
                    firstCellReference = firstCell.CellReference;
                    LastCellReference = lastCell.CellReference;


                    MergeCells mergeCells = new MergeCells();
                    if (worksheetPart.Worksheet.Elements<MergeCells>().Count() > 0)
                    {
                        mergeCells = worksheetPart.Worksheet.Elements<MergeCells>().First();
                    }
                    else
                    {
                        mergeCells = new MergeCells();
                        // Insert a MergeCells object into the specified position.
                        if (worksheetPart.Worksheet.Elements<CustomSheetView>().Count() > 0)
                        {
                            worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<CustomSheetView>().First());
                        }
                        else if (worksheetPart.Worksheet.Elements<DataConsolidate>().Count() > 0)
                        {
                            worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<DataConsolidate>().First());
                        }
                        else if (worksheetPart.Worksheet.Elements<SortState>().Count() > 0)
                        {
                            worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<SortState>().First());
                        }
                        else if (worksheetPart.Worksheet.Elements<AutoFilter>().Count() > 0)
                        {
                            worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<AutoFilter>().First());
                        }
                        else if (worksheetPart.Worksheet.Elements<Scenarios>().Count() > 0)
                        {
                            worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<Scenarios>().First());
                        }
                        else if (worksheetPart.Worksheet.Elements<ProtectedRanges>().Count() > 0)
                        {
                            worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<ProtectedRanges>().First());
                        }
                        else if (worksheetPart.Worksheet.Elements<SheetProtection>().Count() > 0)
                        {
                            worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<SheetProtection>().First());
                        }
                        else if (worksheetPart.Worksheet.Elements<SheetCalculationProperties>().Count() > 0)
                        {
                            worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<SheetCalculationProperties>().First());
                        }
                        else
                        {
                            worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<SheetData>().First());
                        }
                    }
                    MergeCell mergeCell = new MergeCell() { Reference = new StringValue(firstCellReference+":"+LastCellReference) };
                    mergeCells.Append(mergeCell);
                    worksheetPart.Worksheet.Save();


                    rowIndex++;

                    //第二行、数据标题
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
