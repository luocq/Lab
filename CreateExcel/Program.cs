using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using NPOI.SS.UserModel;

namespace CreateExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            string Dir = Environment.CurrentDirectory;
            string fileName =  "Test.xlsx";
            string filePath = Path.Combine(Dir,fileName);
            CreateExcelByNpoi(filePath,GetTestData());
            Console.WriteLine(string.Format("在目录{0}下创建了文件{1}", Dir, fileName));
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
            HeaderRowStyle.FillBackgroundColor = IndexedColors.Red.Index;
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

        /// <summary>
        /// using OpenXmlSDK to CreateExcel 
        /// </summary>
        private static void CreateExcelByOpenXmlSDK(string filePath)
        {
            //string filePath="";
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
            dc = dt.Columns.Add("Money", Type.GetType("System.Decimal"));

            for (int i = 0; i < 100; i++)
            {
                DataRow dr = dt.NewRow();
                dr["name"] = string.Format("姓名_{0}", i);
                dr["Birth"] = DateTime.Now.AddDays(i);
                dr["Count"] = i;
                dr["Money"] = Math.Round((decimal)i / 100, 2);
                dt.Rows.Add(dr);
            }
            return dt;

            
        }


    }   
}
