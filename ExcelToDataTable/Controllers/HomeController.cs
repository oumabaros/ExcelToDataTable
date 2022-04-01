using ExcelToDataTable.Models;
using GemBox.Spreadsheet;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelToDataTable.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private IWebHostEnvironment Environment;
        private IConfiguration Configuration;

        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment _environment,
            IConfiguration _configuration)
        {
            _logger = logger;
            Environment = _environment;
            Configuration = _configuration;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        [HttpPost]
        public async Task<IActionResult> ImportExcel(IFormFile postedFile)
        {
            //Create a Folder.
            string path = Path.Combine(this.Environment.WebRootPath, "FileImports");
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            //Save the uploaded Excel file.
            string fileName = Path.GetFileName(postedFile.FileName);
            string filePath = Path.Combine(path, fileName);
            using (FileStream stream = new FileStream(filePath, FileMode.Create))
            {
                postedFile.CopyTo(stream);
            }

            SpreadsheetInfo.SetLicense("EDWH-6KJO-D7SA-92EZ");

            var workbook = ExcelFile.Load(filePath);
            //Get first worksheet
            var ws = workbook.Worksheets[0];
            
            //create Datatable using worksheet data
            DataTable dt =ws.CreateDataTable(new CreateDataTableOptions()
            {
                ColumnHeaders = true,
                StartRow = 0,
                NumberOfColumns = ws.CalculateMaxUsedColumns(),
                NumberOfRows = ws.Rows.Count,
            });

            return Redirect("~/Home/Index");
        }
       
    }
}
