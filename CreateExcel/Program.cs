using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using NPOI.SS.UserModel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;
using System.Reflection;


namespace CreateExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            string Dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string fileName = "Test.xlsx";
            string filePath = Path.Combine(Dir, fileName);


            Console.WriteLine("输入1使用NPOI输出Excel,输入2使用OpenXML输出Excel");

            try
            {
                int i = int.Parse(Console.ReadLine());
                if (i == 1)
                {
                    CreateExcelByNpoi(filePath, GetTestData());
                    Console.WriteLine(string.Format("在目录{0}下创建了文件{1}", Dir, fileName));
                }
                else if (i == 2)
                {
                    CreateExcelByOpenXmlSDK(filePath, GetTestData());
                    Console.WriteLine(string.Format("在目录{0}下创建了文件{1}", Dir, fileName));
                }
                else
                {
                    Console.WriteLine("输入不正确");
                 
                }
                
            }
            catch (Exception ex)
            {
                Console.WriteLine("发生异常");
            }                              
            Console.ReadKey();
        }

        /// <summary>
        /// using npoi to create new excel document
        /// </summary>
        /// <param name="filePath"></param>
        private static void CreateExcelByNpoi(string filePath,DataTable dt)
        {
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
            IWorkbook workbook = null;
            workbook = new NPOI.XSSF.UserModel.XSSFWorkbook();

            //创建格式
            IDataFormat format = workbook.CreateDataFormat();

            ICellStyle HeaderRowStyle = workbook.CreateCellStyle();//第一行
            ICellStyle DateStyle = workbook.CreateCellStyle();  //日期
            ICellStyle MoneyStyle = workbook.CreateCellStyle(); //金额
            ICellStyle PercentStyle = workbook.CreateCellStyle();//百分比
            //HeaderRowStyle.FillForegroundColor = IndexedColors.Red.Index;
            HeaderRowStyle.FillBackgroundColor = NPOI.SS.UserModel.IndexedColors.Red.Index;
            HeaderRowStyle.FillPattern = FillPattern.BigSpots;
            DateStyle.DataFormat = format.GetFormat("yyyy-MM-dd");
            MoneyStyle.DataFormat = 0x2c;
            PercentStyle.DataFormat = 9;


            ISheet sheet = workbook.CreateSheet("sheetName");

            IRow HeadRow = sheet.CreateRow(0);            
            
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ICell cell = HeadRow.CreateCell(i);
                cell.SetCellValue(dt.Columns[i].ColumnName);
            }

            HeadRow.RowStyle = HeaderRowStyle;
            for (int i = 0; i < dt.Rows.Count; i++)
            {

                IRow row = sheet.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    string value = dt.Rows[i][j].ToString();
                    ICell cell = row.CreateCell(j);
                    String columnType = dt.Columns[j].DataType.ToString();
                    switch (columnType)
                    {
                        case "System.String":
                            cell.SetCellValue(value);
                            break;
                        case "System.Boolean"://布尔型
                            bool boolV = false;
                            bool.TryParse(value, out boolV);
                            cell.SetCellValue(boolV);
                            break;
                        case "System.Int16":
                        case "System.Int64":
                        case "System.Byte":
                        case "System.Int32":
                            int intV = 0;
                            int.TryParse(value, out intV);
                            cell.SetCellValue(intV);
                            cell.CellStyle = MoneyStyle;
                            break;
                        case "System.DateTime":
                            DateTime t_date;
                            DateTime.TryParse(value, out t_date);
                            cell.SetCellValue(t_date);
                            cell.CellStyle = DateStyle;
                            break;
                        case "System.Decimal":
                        case "System.Double":
                            Double t_dec;
                            Double.TryParse(value, out t_dec);
                            cell.SetCellValue(t_dec);
                            cell.CellStyle = PercentStyle;
                            break;
                        case "System.DBNull"://空值处理
                            cell.SetCellValue("");
                            break;
                        default:
                            break;
                    }                            
                }
            }         
            FileStream sw = File.Create(filePath);
            workbook.Write(sw);
            sw.Close();
        }

        private static void CreateExcel(string filePath, List<string> tmp)
        {
 
        }


        /// <summary>
        /// using OpenXmlSDK to CreateExcel 
        /// </summary>
        private static void CreateExcelByOpenXmlSDK(string filePath,DataTable dt)
        {
            //string filePath="";
            MemoryStream stream = new MemoryStream();
            SpreadsheetDocument Doc = SpreadsheetDocument.Create(filePath, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
            Doc.PackageProperties.Creator = "PKUSCE";
            Doc.PackageProperties.Created = DateTime.Now;
            Doc.PackageProperties.LastModifiedBy = "PKUSCE";
            


            using (Doc)
            {
               //创建WorkbookPart，在代码中主要使用这个相当于xml的root elements， spreadSheet.AddWorkbookPart()， 虽然是"Add"方法， 但你只能加一个。
                Doc.AddWorkbookPart();

                WorkbookStylesPart workbookStylesPart1 = Doc.WorkbookPart.AddNewPart<WorkbookStylesPart>("rId3");
                GenerateWorkbookStylesPart1Content(workbookStylesPart1);
                
                Doc.WorkbookPart.Workbook = new Workbook();
              

                //添加工作表(Sheet)
                WorksheetPart worksheetPart = InsertWorksheet(Doc.WorkbookPart,"测试");

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
                                break;
                            case "System.DateTime":
                                cell.DataType = new EnumValue<CellValues>(CellValues.Date);
                                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue { Text = DateTime.Now.ToString("yyyy-MM-dd") };
                                cell.StyleIndex = 1;
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
        /// 
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

        //http://www.cnblogs.com/qimuxiaokai/articles/2864947.html
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

            CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)2U };

            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            Alignment alignment2 = new Alignment() { Vertical = VerticalAlignmentValues.Center };

            cellFormat2.Append(alignment2);

            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)14U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true };
            Alignment alignment3 = new Alignment() { Vertical = VerticalAlignmentValues.Center };

            cellFormat3.Append(alignment3);

            cellFormats1.Append(cellFormat2);
            cellFormats1.Append(cellFormat3);

            CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)1U };
            CellStyle cellStyle1 = new CellStyle() { Name = "常规", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

            cellStyles1.Append(cellStyle1);
            DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)0U };
            TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };


            stylesheet.Append(fonts);
            stylesheet.Append(fills);
            stylesheet.Append(borders);
            stylesheet.Append(cellStyleFormats);
            stylesheet.Append(cellFormats1);
            stylesheet.Append(cellStyles1);
            stylesheet.Append(differentialFormats1);
            stylesheet.Append(tableStyles1);
            //stylesheet.Append(stylesheetExtensionList1);

            workbookStylesPart.Stylesheet = stylesheet;
        }


        /// <summary>
        /// 准备测试数据
        /// </summary>
        /// <returns></returns>
        private static DataTable GetTestData()
        {
            DataTable dt = new DataTable();
            DataColumn dc = null;

            //DataRow dr= dt.NewRow();
            
            
            dc = dt.Columns.Add("name", Type.GetType("System.String"));
            dc = dt.Columns.Add("Birth", Type.GetType("System.DateTime"));
            dc = dt.Columns.Add("Count", Type.GetType("System.Int32"));
            dc = dt.Columns.Add("Money", Type.GetType("System.Double"));

            for (int i = 0; i < 100; i++)
            {
                DataRow dr = dt.NewRow();
                dr["name"] = string.Format("姓名_{0}", i);
                dr["Birth"] = DateTime.Now.AddDays(i);
                dr["Count"] = i;
                dr["Money"] = Math.Round((double)i / 100, 2);
                dt.Rows.Add(dr);
            }
            return dt;

            
        }


    }   
}
