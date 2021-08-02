using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using HtmlAgilityPack;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WebAppHtmlToExcel.Models;
using OfficeOpenXml;
using System.Drawing;
using OfficeOpenXml.Style;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeOpenXml.Drawing.Chart;

namespace WebAppHtmlToExcel.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        public IWebHostEnvironment hostingEnvironment { get; private set; }

        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment hostingEnvironment)
        {
            _logger = logger;
            this.hostingEnvironment = hostingEnvironment;
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
        public FileResult WriteDataToExcel()
        {
            
            using (var libro = new ExcelPackage())
            {
                string excelContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

                //create a WorkSheet
                ExcelWorksheet worksheet = libro.Workbook.Worksheets.Add("Hoja 1");

                //fill cell data with a loop, note that row and column indexes start at 1
                Random rnd = new Random();
                for (int i = 1; i <= 10; i++)
                {
                    worksheet.Cells[1, i].Value = "Valor " + i;
                    worksheet.Cells[2, i].Value = rnd.Next(5, 15);
                }

                //create a new piechart of type Pie3D
                ExcelPieChart pieChart = worksheet.Drawings.AddChart("pieChart", eChartType.Pie3D) as ExcelPieChart;

                //set the title
                pieChart.Title.Text = "Ejemplo: Grafica de Torta";

                //select the ranges for the pie. First the values, then the header range
                pieChart.Series.Add(ExcelRange.GetAddress(2, 1, 2, 10), ExcelRange.GetAddress(1, 1, 1, 10));

                //position of the legend
                pieChart.Legend.Position = eLegendPosition.Bottom;

                //show the percentages in the pie
                pieChart.DataLabel.ShowPercent = true;

                //size of the chart
                pieChart.SetSize(500, 400);

                //add the chart at cell C5
                pieChart.SetPosition(4, 0, 2, 0);


                //create a new BarChart of type BarClustered
                ExcelBarChart barChart = worksheet.Drawings.AddChart("barChart", eChartType.BarClustered) as ExcelBarChart;

                //set the title
                barChart.Title.Text = "Ejemplo: Grafica de Barra";

                //select the ranges for the bar. First the values, then the header range
                barChart.Series.Add(ExcelRange.GetAddress(2, 1, 2, 10), ExcelRange.GetAddress(1, 1, 1, 10));

                //position of the legend
                barChart.Legend.Position = eLegendPosition.Bottom;


                //size of the chart
                barChart.SetSize(500, 400);

                //add the chart at cell K5
                barChart.SetPosition(4, 0, 10, 0);


                return File(libro.GetAsByteArray(), excelContentType, "EjemplosCharts.xlsx");
            }


        }


        //private static Stream GetStreamFromUrl(string url)
        //{
        //    byte[] imageData = null;

        //    using (var wc = new System.Net.WebClient())
        //    {
        //        imageData = wc.DownloadData(url);
        //    }

        //    return new MemoryStream(imageData);
        //}


        //public JsonResult agilpack()
        //{

        //    //string url = "localhost:45220.htm";// string.Format("{0}{1}",Request.Host.Value, Request.Path) ;
        //    string url = "http://localhost:45220";


        //    var doc = new HtmlWeb().Load(url);
        //    List<Tuple<string, string>> tupla = new List<Tuple<string, string>>();
        //    var HTMLTableTRList = from table in doc.DocumentNode.SelectNodes("//table").Where(x => x.Attributes["id"].Value == "header-table").Cast<HtmlNode>()
        //                          from row in table.SelectNodes("tr").Cast<HtmlNode>()
        //                          from cell in row.SelectNodes("th|td").Cast<HtmlNode>()
        //                          select new { Table_Name = table.Id, Cell_Text = cell.InnerText };

        //    // now showing output of parsed HTML table
        //    foreach (var cell in HTMLTableTRList)
        //    {
        //        tupla.Add(new Tuple<string, string>(cell.Table_Name, cell.Cell_Text));
        //    }

            


        //    //var tabla1 = from table in document.DocumentNode.SelectNodes("table").Where(x => x.Attributes["id"].Value == "header-table").Cast<HtmlNode>()
        //    //             from row in table.SelectNodes("tr").Cast<HtmlNode>()
        //    //             from cell in row.SelectNodes("th|td").Cast<HtmlNode>()
        //    //             select new
        //    //             {
        //    //                 Table_Name = table.Id == null ? "" : table.Name,
        //    //                 Cell_Text = cell.InnerText == null ? "" : cell.InnerText
        //    //             };




        //    return Json(tupla);

        //}

        //[HttpPost]
        //public FileResult WriteDataToExcel()
        //{

        //    //string url = "localhost:45220.htm";// string.Format("{0}{1}",Request.Host.Value, Request.Path) ;
        //    string url = "http://localhost:45220";
        //    var document = new HtmlWeb().Load(url);

        //    var data = from tr in document.DocumentNode.Descendants("img")
        //               .Where(x => x.Attributes["name"].Value == "dragon")
        //               .Select(e => e.GetAttributeValue("src", null))
        //               select tr;

        //    string imagePath = data.FirstOrDefault().ToString(); //"https://es.web.img2.acsta.net/r_654_368/newsv7/21/05/20/19/43/3423101.jpg";

        //    //var data = from tr in document.DocumentNode.Descendants("img")
        //    //            from td in tr.Descendants("td").Where(x => x.Attributes["class"].Value == "name")
        //    //            where td.InnerText.Trim() == "Test1"
        //    //        select tr;


        //    DataTable dt = getData();
        //    //Name of File  
        //    string fileName = "Sample.xlsx";
        //    using (XLWorkbook wb = new XLWorkbook())
        //    {



        //        var ws = wb.Worksheets.Add("Sheet1");
        //        ws.Column("A").Width = 93;
        //        ws.Row(1).Height = 270;
        //        var image = ws.AddPicture(GetStreamFromUrl(imagePath));
        //        image.MoveTo(ws.Cell("A1"));
        //        image.Name = "logo";


        //        using (MemoryStream stream = new MemoryStream())
        //        {
        //            wb.SaveAs(stream);
        //            //Return xlsx Excel File  
        //            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        //        }
        //    }
        //}


    }
}
