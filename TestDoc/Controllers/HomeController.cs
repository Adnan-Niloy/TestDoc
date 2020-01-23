using Microsoft.Office.Interop.Word;
using System.Web.Mvc;
using TestDoc.Models;

namespace TestDoc.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
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

        public void ExportToWord(DocProperties model)
        {
            var savePath = Server.MapPath("~/Docs/PersonInfo.docx");
            var templatePath = Server.MapPath("~/Docs/Template.docx");
            var app = new Application();
            var doc = app.Documents.Open(templatePath);
            doc.Activate();

            if (doc.Bookmarks.Exists("RNO"))
            {
                doc.Bookmarks["RNO"].Range.Text = model.Rno;
            }
            if (doc.Bookmarks.Exists("Code"))
            {
                doc.Bookmarks["Code"].Range.Text = model.Code;
            }
            if (doc.Bookmarks.Exists("Description"))
            {
                doc.Bookmarks["Description"].Range.Text = model.Description;
            }

            doc.SaveAs2(savePath);
            doc.Close();

            //byte[] fileBytes = System.IO.File.ReadAllBytes(savePath);
            //string fileName = "MyFile.docx";
            //return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, fileName);

            Response.ContentType = "application/docx";
            Response.AppendHeader("Content-Disposition", "attachment; filename=SailBig.docx");
            Response.TransmitFile(savePath);
            Response.End();

            if (System.IO.File.Exists(savePath))
            {
                System.IO.File.Delete(savePath);
            }
        }
    }
}