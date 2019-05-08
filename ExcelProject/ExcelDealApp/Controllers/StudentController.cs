using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using ExcelDealApp.Models;

namespace ExcelDealApp.Controllers
{
    public class StudentController : Controller
    {
        // GET: Student
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Import(HttpPostedFileBase excelfile) {
            if (excelfile == null || excelfile.ContentLength == 0)
            {
                ViewBag.Error = "請選擇一個excel檔案! <br />";
                return View("Index");
            }
            else {
                if (excelfile.FileName.EndsWith("xls") || excelfile.FileName.EndsWith("xlsx"))
                {
                    string path = Server.MapPath("~/Content/" + excelfile.FileName);
                    if (System.IO.File.Exists(path))
                        System.IO.File.Delete(path);
                    excelfile.SaveAs(path);
                    // Read Data from excel file
                    Excel.Application application = new Excel.Application();
                    Excel.Workbook workbook = application.Workbooks.Open(path);
                    Excel.Worksheet worksheet = workbook.ActiveSheet;
                    Excel.Range range = worksheet.UsedRange;
                    List<AttendRecord> listAttRecords = new List<AttendRecord>();
                    for (int row = 2; row <= range.Rows.Count; row++)
                    {
                        AttendRecord r = new AttendRecord();
                        r.SID = ((Excel.Range)range.Cells[row, 1]).Text;
                        r.Years = ((Excel.Range)range.Cells[row, 2]).Text;
                        r.Term = ((Excel.Range)range.Cells[row, 3]).Text;
                        r.VacType = ((Excel.Range)range.Cells[row, 4]).Text;
                        r.ClassCounts = ((Excel.Range)range.Cells[row, 5]).Text;
                        listAttRecords.Add(r);
                    }
                    ViewBag.ListAttendRecord = listAttRecords;
                    return View("Success");
                }
                else {
                    ViewBag.Error = "檔案型態不正確! <br/>";
                    return View("Index");
                }
            }            
          
        }
    }
}