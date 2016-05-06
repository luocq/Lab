using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;

namespace CreateExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            string Dir = Environment.CurrentDirectory;
            string filePath = Path.Combine(Dir, "Test.xlsx");
            CreateExcelByNpoi(filePath);
        }

        /// <summary>
        /// using npoi to create new excel document
        /// </summary>
        /// <param name="filePath"></param>
        private static void CreateExcelByNpoi(string filePath)
        {
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
            IWorkbook workbook = null;
            workbook = new NPOI.XSSF.UserModel.XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("sheetName");
            for (int i = 0; i < 10; i++)
            {
                IRow row = sheet.CreateRow(i);
                for (int j = 0; j < 10; j++)
                {
                    ICell cell = row.CreateCell(j);
                    cell.SetCellValue(string.Format("{0}:{1}", i, j));
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

        }


    }   
}
