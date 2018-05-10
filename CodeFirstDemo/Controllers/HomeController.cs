using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace CodeFirstDemo.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            using (DAL.CodeFirstDemoEntities db = new DAL.CodeFirstDemoEntities())
            {
                for (int co = 0; co < 3; co++)
                {
                    Models.Project p = new Models.Project();
                    p.Name = string.Format("项目{0}", co.ToString());
                    db.Projects.Add(p);
                }
                db.SaveChanges();
                int i = db.Projects.Count();
                
            }
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}