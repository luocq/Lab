using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelHelper;

namespace ExportExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            DataTable dt = GetTestData();
            //ExcelHelper.ExcelHelper.CreateExcelFromDataTable(dt);
            //Console.ReadKey();            
        }

        /// <summary>
        /// 准备测试数据
        /// </summary>
        /// <returns></returns>
        private static DataTable GetTestData()
        {
            DataTable dt = new DataTable("test data");
            DataColumn dc = null;

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
