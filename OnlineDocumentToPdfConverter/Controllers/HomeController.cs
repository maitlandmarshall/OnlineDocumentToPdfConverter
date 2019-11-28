using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;

namespace OnlineDocumentToPdfConverter.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IWebHostEnvironment _hosting;

        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment hosting)
        {
            _logger = logger;
            _hosting = hosting;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public async Task<FileStreamResult> Index(IFormFile file)
        {
            if (file == null)
                throw new ArgumentNullException(nameof(file));

            string fileName = file.FileName;

            if (Path.GetExtension(fileName) != ".docm")
                throw new ArgumentException("Can only convert docm files to pdf.", nameof(file));

            Guid fileGuid = Guid.NewGuid();
            string conversionFolderPath = Path.Combine(this._hosting.WebRootPath, "conversions");

            if (!Directory.Exists(conversionFolderPath))
                Directory.CreateDirectory(conversionFolderPath);

            string fileSystemName = Path.Combine(conversionFolderPath, $"{fileGuid}.docm");
            string fileSystemPdfName = Path.Combine(conversionFolderPath, $"{fileGuid}.pdf");

            // Write the file to the filesystem
            using (FileStream fs = new FileStream(fileSystemName, FileMode.CreateNew, FileAccess.Write))
            {
                await file.CopyToAsync(fs);
            }

            Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office.Interop.Word.Application();
            appWord.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;

            Microsoft.Office.Interop.Word.Document wordDocument = appWord.Documents.Open(fileSystemName, ReadOnly: true);
            wordDocument.ExportAsFixedFormat(fileSystemPdfName, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);
            wordDocument.Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges);
            appWord.Quit();

            FileStream pdfFileStream = new FileStream(fileSystemPdfName, FileMode.Open, FileAccess.Read);
            return new FileStreamResult(pdfFileStream, "application/pdf");
        }
    }
}
