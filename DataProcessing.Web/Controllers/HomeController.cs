using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using System.Web.UI.WebControls;
using DataProcessing.Web.Models;

namespace DataProcessing.Web.Controllers
{
    public class HomeController : Controller
    {
        // GET: Home
        public ActionResult Index()
        {
            var model = new HomeViewModel();
            return View(model);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        // GET: Home/Create
        public async Task<ActionResult> Convert(HttpPostedFileBase file)
        {
            var model = new HomeViewModel();
            if (file != null && file.ContentLength > 0 && file.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            {
                var excelCsvConverter = new ExcelCsvDataConverter();
                try
                {
                    using (var csvStream = excelCsvConverter.Convert(file.InputStream))
                    {
                        using (var memoryStream = new MemoryStream())
                        {
                            await csvStream.CopyToAsync(memoryStream);
                            return File(memoryStream.ToArray(), "text/csv", String.Format("{0}.{1}",
                                Path.GetFileNameWithoutExtension(file.FileName),
                                "csv"));
                        }
                    }
                }
                catch (Exception)
                {
                    model.UploadErrorMessage = "Unfortunately, we were unable to convert the file.";
                }
            }
            else
            {
                model.UploadErrorMessage = "Only non-empty Excel (.xlsx) files can be uploaded.";
            }
            return View("Index", model);
        }
    }
}
