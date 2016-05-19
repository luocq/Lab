﻿using System;
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
    class ExcelHelper
    {      
        public void CreateExcel(DataTable dt)
        {
         //   Assert.isTrue(false);

            MemoryStream stream = new MemoryStream();
            SpreadsheetDocument Doc = SpreadsheetDocument.Create(stream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
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
                WorksheetPart worksheetPart = InsertWorksheet(Doc.WorkbookPart, "测试");

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

            //StylesheetExtensionList stylesheetExtensionList1 = new StylesheetExtensionList();

            //StylesheetExtension stylesheetExtension1 = new StylesheetExtension() { Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" };
            //stylesheetExtension1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            //X14.SlicerStyles slicerStyles1 = new X14.SlicerStyles() { DefaultSlicerStyle = "SlicerStyleLight1" };

            //stylesheetExtension1.Append(slicerStyles1);

            //StylesheetExtension stylesheetExtension2 = new StylesheetExtension() { Uri = "{9260A510-F301-46a8-8635-F512D64BE5F5}" };
            //stylesheetExtension2.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            //X15.TimelineStyles timelineStyles1 = new X15.TimelineStyles() { DefaultTimelineStyle = "TimeSlicerStyleLight1" };

            //stylesheetExtension2.Append(timelineStyles1);

            //stylesheetExtensionList1.Append(stylesheetExtension1);
            //stylesheetExtensionList1.Append(stylesheetExtension2);

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
    }

    public class SpreadsheetStyle : ICloneable
    {
        //Constructors
        internal SpreadsheetStyle(Font font)
        {
            AddFont(font);
        }

        internal SpreadsheetStyle(Fill fill)
        {
            AddFill(fill);
        }

        internal SpreadsheetStyle(Border border)
        {
            AddBorder(border);
        }

        internal SpreadsheetStyle(Alignment alignment)
        {
            AddAlignment(alignment);
        }

        internal SpreadsheetStyle(NumberingFormat format)
        {
            AddFormat(format);
        }

        internal SpreadsheetStyle(Font font, Fill fill, Border border, Alignment alignment, NumberingFormat format)
        {
            AddFont(font);
            AddFill(fill);
            AddBorder(border);
            AddAlignment(alignment);
            AddFormat(format);
        }

        //Properties
        protected internal Italic Italic { get; set; }
        protected internal Bold Bold { get; set; }
        protected internal Underline Underline { get; set; }
        protected internal Color Color { get; set; }
        protected internal FontSize FontSize { get; set; }
        protected internal FontName FontName { get; set; }
        protected internal FontFamily FontFamily { get; set; }
        protected internal FontScheme FontScheme { get; set; }

        protected internal PatternFill PatternFill { get; set; }

        protected internal TopBorder TopBorder { get; set; }
        protected internal LeftBorder LeftBorder { get; set; }
        protected internal BottomBorder BottomBorder { get; set; }
        protected internal RightBorder RightBorder { get; set; }

        protected internal BooleanValue WrapText { get; set; }
        protected internal EnumValue<VerticalAlignmentValues> VerticalAlignment { get; set; }
        protected internal EnumValue<HorizontalAlignmentValues> HorizontalAlignment { get; set; }

        public StringValue FormatCode;

        ///<summary>
        ///Sets or gets whether the style is italic
        ///</summary>
        public bool IsItalic
        {
            get
            {
                return Italic != null;
            }
            set
            {
                if (value)
                {
                    if (Italic == null) Italic = new Italic();
                }
                else
                {
                    if (Italic != null) Italic = null;
                }
            }
        }

        ///<summary>
        ///Sets or gets whether the style is bold
        ///</summary>
        public bool IsBold
        {
            get
            {
                return Bold != null;
            }
            set
            {
                if ((value))
                {
                    if (Bold == null) Bold = new Bold();
                }
                else
                {
                    if (Bold != null) Bold = null;
                }
            }
        }

        ///<summary>
        ///Sets or gets whether the style is underline
        ///</summary>
        public bool IsUnderline
        {
            get
            {
                return Underline != null;
            }
            set
            {
                if ((value))
                {
                    if (Underline == null) Underline = new Underline();
                }
                else
                {
                    if (Underline != null) Underline = null;
                }
            }
        }

        ///<summary>
        ///Sets or gets whether the style is underline
        ///</summary>
        public bool IsWrapped
        {
            get
            {
                if (WrapText == null || !WrapText.HasValue) return false;
                return WrapText.Value;
            }
            set
            {
                if (WrapText == null) WrapText = new BooleanValue();
                WrapText.Value = value;
            }
        }

        /// <summary>
        /// Returns the default SpreadsheetStyle obejct for the spreadsheet provided.
        /// </summary>
        public static SpreadsheetStyle GetDefault(SpreadsheetDocument spreadsheet)
        {
            return SpreadsheetReader.GetDefaultStyle(spreadsheet);
        }

        ///<summary>
        ///Sets the color using R,G and B hex values eg "FF0000"
        ///</summary>
        public void SetColor(string rgb)
        {
            Color.Theme = null;
            Color.Rgb = "FF" + rgb;
        }

        ///<summary>
        ///Sets the color using R,G and B hex values eg "FF0000"
        ///</summary>
        public void SetBackgroundColor(string rgb)
        {
            if (PatternFill.BackgroundColor == null) PatternFill.BackgroundColor = new BackgroundColor();
            PatternFill.BackgroundColor.Theme = null;
            PatternFill.PatternType = new EnumValue<PatternValues>(PatternValues.Solid);
            PatternFill.BackgroundColor.Rgb = "FF" + rgb;

            if (PatternFill.ForegroundColor == null) PatternFill.ForegroundColor = new ForegroundColor();
            PatternFill.ForegroundColor.Theme = null;
            PatternFill.ForegroundColor.Rgb = "FF" + rgb;
        }

        ///<summary>
        ///Sets all four border color and styles.
        ///</summary>
        public void SetBorder(string rgb, BorderStyleValues style)
        {
            SetBorder(TopBorder, rgb, style);
            SetBorder(LeftBorder, rgb, style);
            SetBorder(RightBorder, rgb, style);
            SetBorder(BottomBorder, rgb, style);
        }

        ///<summary>
        ///Sets the top border's color and style.
        ///</summary>
        public void SetBorderTop(string rgb, BorderStyleValues style)
        {
            SetBorder(TopBorder, rgb, style);
        }

        ///<summary>
        ///Sets the left border's color and style.
        ///</summary>
        public void SetBorderLeft(string rgb, BorderStyleValues style)
        {
            SetBorder(LeftBorder, rgb, style);
        }

        ///<summary>
        ///Sets the bottom border's color and style.
        ///</summary>
        public void SetBorderBottom(string rgb, BorderStyleValues style)
        {
            SetBorder(BottomBorder, rgb, style);
        }

        ///<summary>
        ///Sets the right border's color and style.
        ///</summary>
        public void SetBorderRight(string rgb, BorderStyleValues style)
        {
            SetBorder(RightBorder, rgb, style);
        }

        ///<summary>
        ///Sets an individual border's color and style.
        ///</summary>
        protected void SetBorder(BorderPropertiesType item, string rgb, BorderStyleValues style)
        {
            if (item.Color == null) item.Color = new Color();
            item.Color.Theme = null;
            item.Color.Rgb = "FF" + rgb;
            item.Style = new EnumValue<BorderStyleValues>(style);
        }

        ///<summary>
        ///Sets the horizontal alignment value
        ///</summary>
        public void SetHorizontalAlignment(HorizontalAlignmentValues value)
        {
            if (HorizontalAlignment == null) HorizontalAlignment = new EnumValue<HorizontalAlignmentValues>();
            HorizontalAlignment.Value = value;
        }

        ///<summary>
        ///Sets the vertical alignment value
        ///</summary>
        public void SetVerticalAlignment(VerticalAlignmentValues value)
        {
            if (VerticalAlignment == null) VerticalAlignment = new EnumValue<VerticalAlignmentValues>();
            VerticalAlignment.Value = value;
        }

        ///<summary>
        ///Sets the format code value
        ///</summary>
        public void SetFormat(string format)
        {
            if (FormatCode == null) FormatCode = new StringValue();
            FormatCode.Value = format;
        }

        ///<summary>
        ///Sets all four border color and styles.
        ///</summary>
        public void ClearBorder()
        {
            ClearBorder(TopBorder);
            ClearBorder(LeftBorder);
            ClearBorder(RightBorder);
            ClearBorder(BottomBorder);
        }

        ///<summary>
        ///Sets the top border's color and style.
        ///</summary>
        public void ClearBorderTop()
        {
            ClearBorder(TopBorder);
        }

        ///<summary>
        ///Sets the left border's color and style.
        ///</summary>
        public void ClearBorderLeft()
        {
            ClearBorder(LeftBorder);
        }

        ///<summary>
        ///Sets the bottom border's color and style.
        ///</summary>
        public void ClearBorderBottom()
        {
            ClearBorder(BottomBorder);
        }

        ///<summary>
        ///Sets the right border's color and style.
        ///</summary>
        public void ClearBorderRight()
        {
            ClearBorder(RightBorder);
        }

        ///<summary>
        ///Sets an individual border's color and style.
        ///</summary>
        protected internal void ClearBorder(BorderPropertiesType item)
        {
            if (item.Color == null) item.Color = new Color();
            item.Color.Theme = null;
            item.Color.Rgb = null;
            item.Style = new EnumValue<BorderStyleValues>(BorderStyleValues.None);
        }

        ///<summary>
        ///Overrides any style information by copying from the Font object provided.
        ///</summary>
        public void AddFont(Font font)
        {
            Italic = null;
            Bold = null;
            Underline = null;

            if (font.ChildElements.OfType<Italic>().Count() > 0) Italic = new Italic();
            if (font.ChildElements.OfType<Bold>().Count() > 0) Bold = new Bold();
            if (font.ChildElements.OfType<Underline>().Count() > 0) Underline = new Underline();

            if (font.ChildElements.OfType<Color>().Count() > 0) Color = font.ChildElements.OfType<Color>().First().CloneElement<Color>();
            if (font.ChildElements.OfType<FontSize>().Count() > 0) FontSize = font.ChildElements.OfType<FontSize>().First().CloneElement<FontSize>();
            if (font.ChildElements.OfType<FontName>().Count() > 0) FontName = font.ChildElements.OfType<FontName>().First().CloneElement<FontName>();
            if (font.ChildElements.OfType<FontFamily>().Count() > 0) FontFamily = font.ChildElements.OfType<FontFamily>().First().CloneElement<FontFamily>();
            if (font.ChildElements.OfType<FontScheme>().Count() > 0) FontScheme = font.ChildElements.OfType<FontScheme>().First().CloneElement<FontScheme>();
        }

        ///<summary>
        ///Overrides any fill style information by copying from from the Fill object provided
        ///</summary>
        public void AddFill(Fill fill)
        {
            PatternFill = fill.ChildElements.OfType<PatternFill>().First().CloneElement<PatternFill>();
        }

        ///<summary>
        ///Overrides any border style information by copying from the Border object provided
        ///</summary>
        public void AddBorder(Border border)
        {
            this.TopBorder = border.TopBorder.CloneElement<TopBorder>();
            this.LeftBorder = border.LeftBorder.CloneElement<LeftBorder>();
            this.BottomBorder = border.BottomBorder.CloneElement<BottomBorder>();
            this.RightBorder = border.RightBorder.CloneElement<RightBorder>();
        }

        ///<summary>
        ///Overrides any style information by copying from the Alignment object provided.
        ///</summary>
        public void AddAlignment(Alignment alignment)
        {
            WrapText = new BooleanValue();
            HorizontalAlignment = new EnumValue<HorizontalAlignmentValues>();
            VerticalAlignment = new EnumValue<VerticalAlignmentValues>();

            if (alignment != null)
            {
                if (alignment.WrapText != null && alignment.WrapText.HasValue) WrapText.Value = alignment.WrapText.Value;
                if (alignment.Horizontal != null && alignment.Horizontal.HasValue) HorizontalAlignment.Value = alignment.Horizontal.Value;
                if (alignment.Vertical != null && alignment.Vertical.HasValue) VerticalAlignment.Value = alignment.Vertical.Value;
            }
        }

        ///<summary>
        ///Overrides any number format information by copying from from the object provided
        ///</summary>
        public void AddFormat(NumberingFormat format)
        {
            FormatCode = new StringValue();
            if (format != null && format.FormatCode != null) FormatCode.Value = format.FormatCode.Value;
        }

        ///<summary>
        ///Returns a new Font object from the style information provided
        ///</summary>
        public Font ToFont()
        {
            Font font = new Font();

            if (Italic != null) font.AppendChild<Italic>(new Italic());
            if (Bold != null) font.AppendChild<Bold>(new Bold());
            if (Underline != null) font.AppendChild<Underline>(new Underline());

            if (Color != null) font.AppendChild<Color>(Color.CloneElement<Color>());
            if (FontSize != null) font.AppendChild<FontSize>(FontSize.CloneElement<FontSize>());
            if (FontName != null) font.AppendChild<FontName>(FontName.CloneElement<FontName>());
            if (FontFamily != null) font.AppendChild<FontFamily>(FontFamily.CloneElement<FontFamily>());
            if (FontScheme != null) font.AppendChild<FontScheme>(FontScheme.CloneElement<FontScheme>());

            return font;
        }

        ///<summary>
        ///Returns a new Font object from the style information provided
        ///</summary>
        public Fill ToFill()
        {
            Fill fill = new Fill();

            fill.AppendChild<PatternFill>(PatternFill.CloneElement<PatternFill>());

            return fill;
        }

        ///<summary>
        ///Returns a new Border object from the style information provided
        ///</summary>
        public Border ToBorder()
        {
            Border border = new Border();

            border.TopBorder = TopBorder.CloneElement<TopBorder>();
            border.LeftBorder = LeftBorder.CloneElement<LeftBorder>();
            border.BottomBorder = BottomBorder.CloneElement<BottomBorder>();
            border.RightBorder = RightBorder.CloneElement<RightBorder>();
            border.DiagonalBorder = new DiagonalBorder();
            return border;
        }

        ///<summary>
        ///Returns a new Alignment object from this object
        ///</summary>
        public Alignment ToAlignment()
        {
            if ((WrapText == null || !WrapText.HasValue) && (HorizontalAlignment == null || !HorizontalAlignment.HasValue) && (VerticalAlignment == null || !VerticalAlignment.HasValue)) return null;

            var alignment = new Alignment();

            if (WrapText != null && WrapText.HasValue) alignment.WrapText = new BooleanValue(WrapText.Value);
            if (HorizontalAlignment != null && HorizontalAlignment.HasValue) alignment.Horizontal = new EnumValue<HorizontalAlignmentValues>(HorizontalAlignment.Value);
            if (VerticalAlignment != null && VerticalAlignment.HasValue) alignment.Vertical = new EnumValue<VerticalAlignmentValues>(VerticalAlignment.Value);

            return alignment;
        }

        ///<summary>
        ///Returns a new NumberFormat object from this object
        ///</summary>
        public NumberingFormat ToNumberFormat()
        {
            if (FormatCode == null || !FormatCode.HasValue) return null;

            NumberingFormat format = new NumberingFormat();
            format.FormatCode = FormatCode;
            return format;
        }

        ///<summary>
        ///Returns a deep copy of this object.
        ///</summary>
        public object Clone()
        {
            return new SpreadsheetStyle(this.ToFont(), this.ToFill(), this.ToBorder(), this.ToAlignment(), this.ToNumberFormat());
        }

        ///<summary>
        ///Determines if the two font settings supplied are the same.
        ///</summary>
        protected internal static bool CompareFont(Font font1, Font font2)
        {
            return CompareFont(new SpreadsheetStyle(font1), new SpreadsheetStyle(font2));
        }

        ///<summary>
        ///Determines if the two font settings supplied are the same.
        ///</summary>
        protected internal static bool CompareFont(Font font1, SpreadsheetStyle font2)
        {
            return CompareFont(new SpreadsheetStyle(font1), font2);
        }

        ///<summary>
        ///Determines if the two font settings supplied are the same.
        ///</summary>
        protected internal static bool CompareFont(SpreadsheetStyle font1, SpreadsheetStyle font2)
        {
            if (!font1.Italic.Compare(font2.Italic)) return false;
            if (!font1.Bold.Compare(font2.Bold)) return false;
            if (!font1.Underline.Compare(font2.Underline)) return false;

            if (!font1.Color.Compare(font2.Color)) return false;
            if (!font1.FontSize.Compare(font2.FontSize)) return false;
            if (!font1.FontName.Compare(font2.FontName)) return false;
            if (!font1.FontFamily.Compare(font2.FontFamily)) return false;
            if (!font1.FontScheme.Compare(font2.FontScheme)) return false;

            return true;
        }

        ///<summary>
        ///Determines if the two fill settings supplied are the same.
        ///</summary>
        protected internal static bool CompareFill(Fill fill1, Fill fill2)
        {
            return CompareFill(new SpreadsheetStyle(fill1), new SpreadsheetStyle(fill2));
        }

        ///<summary>
        ///Determines if the two fill settings supplied are the same.
        ///</summary>
        protected internal static bool CompareFill(Fill fill1, SpreadsheetStyle fill2)
        {
            return CompareFill(new SpreadsheetStyle(fill1), fill2);
        }

        ///<summary>
        ///Determines if the two fill settings supplied are the same.
        ///</summary>
        protected internal static bool CompareFill(SpreadsheetStyle fill1, SpreadsheetStyle fill2)
        {
            if (!fill1.PatternFill.ForegroundColor.Compare(fill2.PatternFill.ForegroundColor)) return false;
            if (!fill1.PatternFill.BackgroundColor.Compare(fill2.PatternFill.BackgroundColor)) return false;
            if (!fill1.PatternFill.PatternType.Compare(fill2.PatternFill.PatternType)) return false;

            return true;
        }

        ///<summary>
        ///Determines if the two border settings supplied are the same.
        ///</summary>
        protected internal static bool CompareBorder(Border border1, Border border2)
        {
            return CompareBorder(new SpreadsheetStyle(border1), new SpreadsheetStyle(border2));
        }

        ///<summary>
        ///Determines if the two border settings supplied are the same.
        ///</summary>
        protected internal static bool CompareBorder(Border border1, SpreadsheetStyle border2)
        {
            return CompareBorder(new SpreadsheetStyle(border1), border2);
        }

        ///<summary>
        ///Determines if the two border settings supplied are the same.
        ///</summary>
        protected internal static bool CompareBorder(SpreadsheetStyle border1, SpreadsheetStyle border2)
        {
            if (!border1.TopBorder.Compare(border2.TopBorder)) return false;
            if (!border1.LeftBorder.Compare(border2.LeftBorder)) return false;
            if (!border1.BottomBorder.Compare(border2.BottomBorder)) return false;
            if (!border1.RightBorder.Compare(border2.RightBorder)) return false;

            return true;
        }

        ///<summary>
        ///Determines if the two alignment settings supplied are the same.
        ///</summary>
        protected internal static bool CompareAlignment(Alignment alignment1, Alignment alignment2)
        {
            return CompareAlignment(new SpreadsheetStyle(alignment1), new SpreadsheetStyle(alignment2));
        }

        ///<summary>
        ///Determines if the two alignment settings supplied are the same.
        ///</summary>
        protected internal static bool CompareAlignment(Alignment alignment1, SpreadsheetStyle alignment2)
        {
            return CompareAlignment(new SpreadsheetStyle(alignment1), alignment2);
        }

        ///<summary>
        ///Determines if the two alignment settings supplied are the same.
        ///</summary>
        protected internal static bool CompareAlignment(SpreadsheetStyle alignment1, SpreadsheetStyle alignment2)
        {
            if (!alignment1.WrapText.Compare(alignment2.WrapText)) return false;
            if (!alignment1.HorizontalAlignment.Compare(alignment2.HorizontalAlignment)) return false;
            if (!alignment1.VerticalAlignment.Compare(alignment2.VerticalAlignment)) return false;

            return true;
        }

        ///<summary>
        ///Determines if the two format settings supplied are the same.
        ///</summary>
        protected internal static bool CompareNumberFormat(NumberingFormat format1, NumberingFormat format2)
        {
            return CompareNumberFormat(new SpreadsheetStyle(format1), new SpreadsheetStyle(format2));
        }

        ///<summary>
        ///Determines if the two format settings supplied are the same.
        ///</summary>
        protected internal static bool CompareNumberFormat(NumberingFormat format1, SpreadsheetStyle format2)
        {
            return CompareNumberFormat(new SpreadsheetStyle(format1), format2);
        }

        ///<summary>
        ///Determines if the two format settings supplied are the same.
        ///</summary>
        protected internal static bool CompareNumberFormat(SpreadsheetStyle format1, SpreadsheetStyle format2)
        {
            if (!format1.FormatCode.Compare(format2.FormatCode)) return false;
            return true;
        }
    }
}