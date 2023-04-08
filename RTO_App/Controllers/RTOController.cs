using Bytescout.Spreadsheet;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using RTO_App.Models;
using System.Diagnostics;
using System.IO;

namespace RTO_App.Controllers
{
    public class RTOController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult Index(RTOForm model)
        {
            Spreadsheet document = new Spreadsheet();
            Bytescout.Spreadsheet.Worksheet sheet = document.Workbook.Worksheets.Add("writeExcelData");
            sheet.Cell("A1").Value = "Employee ID"; 
            sheet.Cell("B1").Value = "Compliance Type"; 
            sheet.Cell("C1").Value = "From"; 
            sheet.Cell("D1").Value = "To";

            sheet.Cell("A2").Value = model.EmployeeID;
            sheet.Cell("B2").Value = model.ComplianceType;
            sheet.Cell("C2").Value = model.From;
            sheet.Cell("D2").Value = model.To;

            if (System.IO.File.Exists(@"C:\Users\Deepa G\OneDrive\Desktop\RTOExcelData\EmployeeOutput.xlsx"))
            {
                System.IO.File.Delete(@"C:\Users\Deepa G\OneDrive\Desktop\RTOExcelData\EmployeeOutput.xlsx");
            }
            document.SaveAs(@"C:\Users\Deepa G\OneDrive\Desktop\RTOExcelData\EmployeeOutput.xlsx");
            document.Close();

            return View("Success");
        }

        public IActionResult Success()
        {
            return View();
        }

    }
}
