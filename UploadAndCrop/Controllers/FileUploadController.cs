using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace UploadAndCrop.Controllers
{
    public class FileUploadController : Controller
    {
        //
        // GET: /FileUpload/
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public JsonResult UploadFile()
        {
            string status = "error";
            int width = 0;
            int height = 0;
            string fileName = "";
            var res = new object();
            try
            {
                HttpPostedFileBase file = Request.Files["img"];
                if (file != null)
                {
                    fileName = Path.GetFileName(file.FileName);// 原始文件名称
                    string fileExtension = Path.GetExtension(fileName); // 文件扩展名
                    string saveName = Guid.NewGuid().ToString() + fileExtension; // 保存文件名称
                    string fileDir = HttpContext.Server.MapPath("~/UploadFiles");//文件存放路径

                    //服务器路径
                    string filepath = System.IO.Path.Combine(fileDir, saveName);

                    file.SaveAs(filepath);
                    using (Image image = Image.FromFile(filepath))
                    {
                        status = "success";
                        width = image.Width;
                        height = image.Height;
                        fileName = saveName;
                    }
                    //var image = System.Drawing.Image.FromFile(filepath);
                    res = new { status = status, width = width, height = height, fileName = fileName };
                    return Json(res, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    return null;
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        public ActionResult CropImage() {
            string imgUrl = Request.Form["imgUrl"].ToString();
            var originalFilePath = Server.MapPath(imgUrl);
            int x = Convert.ToInt32(Math.Round(decimal.Parse(Request.Form["x"].ToString())));
            int y = Convert.ToInt32(Math.Round(decimal.Parse(Request.Form["y"].ToString())));
            int w = Convert.ToInt32(Math.Round(decimal.Parse(Request.Form["w"].ToString())));
            int h = Convert.ToInt32(Math.Round(decimal.Parse(Request.Form["h"].ToString())));

            Bitmap OriginalImage = new Bitmap(originalFilePath);
            int OriginalWidth = OriginalImage.Width;
            int OriginalHeight = OriginalImage.Height;
            Bitmap TargetImage = new Bitmap(w, h);
            using (var g = Graphics.FromImage(TargetImage))
            {
                g.DrawImage(OriginalImage, new Rectangle(0, 0, w, h), new Rectangle(x, y, w, h), GraphicsUnit.Pixel);
            }
            string fileName = Path.GetFileName(originalFilePath);
            string SavePath = Path.Combine(Server.MapPath("/Avatar"), fileName);
            TargetImage.Save(SavePath);
            ViewBag.src = "/Avatar/" + fileName;
            return View();
        }       
    }
}