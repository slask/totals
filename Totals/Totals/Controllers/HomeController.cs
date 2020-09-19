using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Totals.Models;

namespace Totals.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private const string XlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost("FileUpload")]
        public IActionResult Upload(IFormFile formFile)
        {
            Dictionary<string, int> totals = new Dictionary<string, int>();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(formFile.OpenReadStream()))
            {
                //Get the first worksheet in the workbook
                ExcelWorksheet srcSheet = package.Workbook.Worksheets[0];

                int skuCol = 1; //the item name column index (1 based)
                int qtySoldCol = 8; //the qty sold coulmn index (1 based)
                int firstProductRow = 10; // the first row that has actual product data
                int totalRows = srcSheet.Dimension.End.Row;

                for (int row = firstProductRow; row < totalRows; row++)
                {
                    var skuValue = srcSheet.Cells[row, skuCol].Value;
                    if (skuValue == null)
                    {
                        continue;
                    }

                    string sku = skuValue.ToString();
                    if (string.IsNullOrEmpty(sku))
                    {
                        continue;
                    }

                    var qtyValue = srcSheet.Cells[row, qtySoldCol].Value;
                    if (qtyValue == null)
                    {
                        continue;
                    }

                    if (!int.TryParse(qtyValue.ToString(), out int qty))
                    {
                        continue;
                    }

                    if (!totals.ContainsKey(sku))
                    {
                        totals.Add(sku, qty);
                    }
                    else
                    {
                        totals[sku] += qty;
                    }
                }
            }

            byte[] reportBytes;
            using (var destPackage = CreateDestExcelPackage())
            {
                var destSheet = destPackage.Workbook.Worksheets[0];

                // set number styling
                var numberformat = "#,##0";
                var dataCellStyleName = "TableNumber";
                var numStyle = destPackage.Workbook.Styles.CreateNamedStyle(dataCellStyleName);
                numStyle.Style.Numberformat.Format = numberformat;

                //Add values
                int rowIndex = 2;
                foreach (var kvp in totals)
                {
                    destSheet.Cells[rowIndex, 1].Value = kvp.Key;
                    destSheet.Cells[rowIndex, 2].Value = kvp.Value;
                    destSheet.Cells[rowIndex, 2].Style.Numberformat.Format = numberformat;
                    rowIndex++;
                }

                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    destSheet.Cells[1, 1, rowIndex, 2].AutoFitColumns();
                }
                
                reportBytes = destPackage.GetAsByteArray();
            }

            return File(reportBytes, XlsxContentType, $"TotaluriPerProduse.xlsx");
        }

        private ExcelPackage CreateDestExcelPackage()
        {
            var package = new ExcelPackage();
            package.Workbook.Properties.Title = "Total per Produs";
            package.Workbook.Properties.Author = "TTH";
            package.Workbook.Properties.Subject = "Total per Produs";


            var worksheet = package.Workbook.Worksheets.Add("Totaluri");

            //First add the headers
            worksheet.Cells[1, 1].Value = "Produs";
            worksheet.Cells[1, 2].Value = "Total";
            worksheet.Cells[1, 1, 1, 2].Style.Font.Bold = true;
            worksheet.Cells[1, 1, 1, 2].Style.Font.UnderLine = true;
            worksheet.Cells[1, 1, 1, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[1, 1, 1, 2].Style.Fill.BackgroundColor.SetColor(Color.Bisque);

            return package;
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}