using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using ExcelHelper;

namespace Test
{
    public partial class index : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btnExportExcel_Click(object sender, EventArgs e)
        {
            ExcelHelper.ExcelHelper helper = new ExcelHelper.ExcelHelper();

            MemoryStream ms = helper.CreateExcel(GetTestData());
            HttpContext.Current.Response.AppendHeader("Content-Disposition", "attachment;filename=test.xlsx");
            HttpContext.Current.Response.BinaryWrite(ms.ToArray());
            HttpContext.Current.Response.End();
            ms.Close();
            ms = null;
        }

        private DataTable GetTestData()
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