using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

using OfficeOpenXml;

namespace MVC___ImportExcel.Controllers
{
    public class HomeController : Controller
    {
        public class clsExcelFileRow
        { // This property order MUST MATCH the imported Excel file header row (by position):
            public string ESITPMID { get; set; }
            public string Customer { get; set; }
            public string Course { get; set; }
            public string Modality { get; set; }
            public int NumberStudents { get; set; }
            public string DeliveryCity { get; set; }
            public string Language { get; set; }
            public string Location { get; set; }
            public DateTime DeliveryWindowStart { get; set; }
            public DateTime DeliveryWindowEnd { get; set; }
            public string PreferedDateNotes { get; set; }
            public string DeliveryChannel { get; set; }
            public bool ReadyToSchedule { get; set; }
        }

        public class clsIndexViewModel
        {
            public List<clsExcelFileRow> lst_clsExcelFileRow { get; private set; }
            public List<string> lst_clsImportErrors {get;private set;}
            public clsIndexViewModel()
            {
                this.lst_clsExcelFileRow = new List<clsExcelFileRow>();
                this.lst_clsImportErrors = new List<string>();
            }
        }

        public class clsExportedExcelFileHeaderRow
        { // These are used to create the exported Excel file header row:
            public string ESITPMID { get; set; }
            public string Customer { get; set; }
            public string Course { get; set; }
            public string Modality { get; set; }
            public string NumberStudents { get; set; }
            public string DeliveryCity { get; set; }
            public string Language { get; set; }
            public string Location { get; set; }
            public string DeliveryWindowStart { get; set; }
            public string DeliveryWindowEnd { get; set; }
            public string PreferedDateNotes { get; set; }
            public string DeliveryChannel { get; set; }
            public string ReadyToSchedule { get; set; }
        }

        public ActionResult Index()
        {
            ViewBag.RowsImported = 0;
            return View();
        }

        [HttpPost]
        public ActionResult Upload(FormCollection formCollection)
        {
            //var lst_clsExcelFileRow = new List<clsExcelFileRow>(); // List of rows to stick into DB
            var lst_strImportErrors = new List<string>(); // List of import errors to show client
            var obj_clsIndexViewModel = new clsIndexViewModel(); // Returns either Imported Data list or Error list to Index View

            if (Request != null)
            {
                HttpPostedFileBase file = Request.Files["UploadedFile"];
                try
                {
                    if ((file != null && file.ContentLength > 0 && !string.IsNullOrEmpty(file.FileName)))
                    {
                        string fileName = file.FileName;
                        string fileContentType = file.ContentType;
                        byte[] filebytes = new byte[file.ContentLength];
                        var data = file.InputStream.Read(filebytes, 0, Convert.ToInt32(file.ContentLength));
                        
                        using (var package = new ExcelPackage(file.InputStream))
                        {
                            var currentSheet = package.Workbook.Worksheets;
                            var workSheet = currentSheet.First();
                            //ExcelWorksheet workSheet = package.Workbook.Worksheets[0];
                            if (workSheet.Dimension.Start != null)
                            {

                                bool bolBadDataWasFound = false;
                                var start = workSheet.Dimension.Start;
                                var end = workSheet.Dimension.End;
                                var noOfColumns = workSheet.Dimension.End.Column;
                                var noOfRows = workSheet.Dimension.End.Row;

                                for (int intCurrentRow = 2; intCurrentRow <= noOfRows; intCurrentRow++)
                                {
                                    if (workSheet.Cells[intCurrentRow, 1].Value == null || workSheet.Cells[intCurrentRow, 1].Value.ToString() == string.Empty) // Exit loop at first blank row
                                    {
                                        noOfRows = intCurrentRow-1;
                                        break;
                                    }

                                    var objExcelFileRow = new clsExcelFileRow();

                                    // Check all STRING fields for nulls & replace with "" if found:
                                    if (workSheet.Cells[intCurrentRow, 1].Value == null) { objExcelFileRow.ESITPMID = ""; } else { objExcelFileRow.ESITPMID = workSheet.Cells[intCurrentRow, 1].Value.ToString(); }
                                    if (workSheet.Cells[intCurrentRow, 2].Value == null) { objExcelFileRow.Customer = ""; } else { objExcelFileRow.Customer = workSheet.Cells[intCurrentRow, 2].Value.ToString(); }
                                    if (workSheet.Cells[intCurrentRow, 3].Value == null) { objExcelFileRow.Course = ""; } else { objExcelFileRow.Course = workSheet.Cells[intCurrentRow, 3].Value.ToString(); }
                                    if (workSheet.Cells[intCurrentRow, 4].Value == null) { objExcelFileRow.Modality = ""; } else { objExcelFileRow.Modality = workSheet.Cells[intCurrentRow, 4].Value.ToString(); }
                                    //if (workSheet.Cells[intCurrentRow, 5].Value == null) { objExcelFileRow.NumberStudents = ""; } else { objExcelFileRow.NumberStudents = workSheet.Cells[intCurrentRow, 5].Value.ToString(); }
                                    if (workSheet.Cells[intCurrentRow, 6].Value == null) { objExcelFileRow.DeliveryCity = ""; } else { objExcelFileRow.DeliveryCity = workSheet.Cells[intCurrentRow, 6].Value.ToString(); }
                                    if (workSheet.Cells[intCurrentRow, 7].Value == null) { objExcelFileRow.Language = ""; } else { objExcelFileRow.Language = workSheet.Cells[intCurrentRow, 7].Value.ToString(); }
                                    if (workSheet.Cells[intCurrentRow, 8].Value == null) { objExcelFileRow.Location = ""; } else { objExcelFileRow.Location = workSheet.Cells[intCurrentRow, 8].Value.ToString(); }
                                    //objExcelFileRow.DeliveryWindowStart = DateTime.Parse(workSheet.Cells[intCurrentRow, 9].Value.ToString());
                                    //objExcelFileRow.DeliveryWindowEnd = DateTime.Parse(workSheet.Cells[intCurrentRow, 10].Value.ToString());
                                    if (workSheet.Cells[intCurrentRow, 11].Value == null) { objExcelFileRow.PreferedDateNotes = ""; } else { objExcelFileRow.PreferedDateNotes = workSheet.Cells[intCurrentRow, 11].Value.ToString(); }
                                    if (workSheet.Cells[intCurrentRow, 12].Value == null) { objExcelFileRow.DeliveryChannel = ""; } else { objExcelFileRow.DeliveryChannel = workSheet.Cells[intCurrentRow, 12].Value.ToString(); }
                                    //if (workSheet.Cells[intCurrentRow, 13].Value == null) { objExcelFileRow.ReadyToSchedule = ""; } else { objExcelFileRow.ReadyToSchedule = workSheet.Cells[intCurrentRow, 13].Value.ToString(); }

                                    // Check all INT fields for nulls, then validate if something is found:
                                    if (workSheet.Cells[intCurrentRow, 5].Value == null || workSheet.Cells[intCurrentRow, 5].Value.ToString().Trim() == "")
                                    {
                                        objExcelFileRow.NumberStudents = 0;
                                    }
                                    else
                                    {
                                        int intStudentCount = 0;
                                        if (int.TryParse(workSheet.Cells[intCurrentRow, 5].Value.ToString().Trim(), out intStudentCount))
                                        { objExcelFileRow.NumberStudents = intStudentCount; }
                                        else
                                        {
                                            obj_clsIndexViewModel.lst_clsImportErrors.Add("Found bad data at Row " + intCurrentRow + ", Column E. A single number is expected (Ex: 25). Found: " + workSheet.Cells[intCurrentRow, 5].Value.ToString() + " instead.");
                                            bolBadDataWasFound = true;
                                        }
                                    }
                                    
                                    // Check all DATE fields for nulls & validate as DateTime if something is found:
                                    DateTime dtDWS;
                                    if (DateTime.TryParse(workSheet.Cells[intCurrentRow, 9].Value.ToString(), out dtDWS))
                                    { objExcelFileRow.DeliveryWindowStart = dtDWS; }
                                    else
                                    {
                                        obj_clsIndexViewModel.lst_clsImportErrors.Add("Found bad data at Row " + intCurrentRow + ", Column I. A valid Date is expected (Ex: 12/10/2021). Found: " + workSheet.Cells[intCurrentRow, 9].Value.ToString() + " instead.");
                                        bolBadDataWasFound = true;
                                    }
                                    
                                    DateTime dtDWE;
                                    if (DateTime.TryParse( workSheet.Cells[intCurrentRow, 10].Value.ToString(), out dtDWE))
                                    { objExcelFileRow.DeliveryWindowEnd = dtDWE; }
                                    else
                                    {
                                        obj_clsIndexViewModel.lst_clsImportErrors.Add("Found bad data at Row " + intCurrentRow + ", Column J. A valid Date is expected (Ex: 12/10/2021). Found: " + workSheet.Cells[intCurrentRow, 10].Value.ToString() + " instead.");
                                        bolBadDataWasFound = true;
                                    }

                                    // Check all BOOLEAN fields for YES/NO, & validate as BOOLEAN if something else is found:
                                    bool bolRTS;
                                    if (bool.TryParse(workSheet.Cells[intCurrentRow, 13].Value.ToString(), out bolRTS)) // True or False?
                                    {
                                        objExcelFileRow.ReadyToSchedule = bolRTS;
                                    }
                                    else if (workSheet.Cells[intCurrentRow, 13].Value.ToString().Equals("yes", StringComparison.OrdinalIgnoreCase) || workSheet.Cells[intCurrentRow, 13].Value.ToString().Equals("y", StringComparison.OrdinalIgnoreCase))
                                    {
                                        objExcelFileRow.ReadyToSchedule = true;
                                    }
                                    else if (workSheet.Cells[intCurrentRow, 13].Value.ToString().Equals("no", StringComparison.OrdinalIgnoreCase) || workSheet.Cells[intCurrentRow, 13].Value.ToString().Equals("n", StringComparison.OrdinalIgnoreCase))
                                    {
                                        objExcelFileRow.ReadyToSchedule = false;
                                    }
                                    else
                                    {
                                        obj_clsIndexViewModel.lst_clsImportErrors.Add("Found bad data at Row " + intCurrentRow + ", Column M. Yes, Y, No, or N is expected (Ex: yes). Found: " + workSheet.Cells[intCurrentRow, 13].Value.ToString() + " instead.");
                                        bolBadDataWasFound = true;
                                    }

                                    // All row fields checked, and, if good, add the row to the row list:
                                    if (!bolBadDataWasFound)
                                    {
                                        // Add row to row list:
                                        obj_clsIndexViewModel.lst_clsExcelFileRow.Add(objExcelFileRow);
                                    }
                                }

                                // If all data rows are good, then show to client:
                                if (!bolBadDataWasFound)
                                {
                                    // Add other vars for display to user:
                                    ViewBag.ImportResult = "Successful!";
                                    ViewBag.RowsImported = noOfRows - 1;
                                    ViewBag.FileName = fileName.ToString();
                                    ViewBag.SomeText = "Some Text!";

                                    // May Export or Save list of rows to DB here:

                                }
                                else // Bad data found...
                                { // Clear the row list of any imported rows before error occurred:
                                    obj_clsIndexViewModel.lst_clsExcelFileRow.Clear();
                                    ViewBag.RowsImported = 0;
                                    ViewBag.FileName = fileName.ToString();
                                    return View("Index", obj_clsIndexViewModel);
                                }
                            }
                        }
                    }
                    else
                    { // Entire file was empty.
                        obj_clsIndexViewModel.lst_clsExcelFileRow.Clear();
                        ViewBag.RowsImported = 0;
                        ViewBag.ImportResult = "FAILED!!! Error: Uploaded Excel file was empty.";
                    }
                }
                catch (Exception e)
                { // Some other error occurred.
                    obj_clsIndexViewModel.lst_clsExcelFileRow.Clear();
                    ViewBag.RowsImported = 0;
                    ViewBag.ImportResult = "FAILED!!! Error: " + e;
                }
            }
            return View("Index", obj_clsIndexViewModel);
        }

        public ActionResult DownloadTemplate()
        {
            //var filename = "ExcellData.xlsx";
            //using (var package = new OfficeOpenXml.ExcelPackage(filename))
            var lst_clsExcelFileRow = new List<clsExcelFileRow>();

            using (var package = new OfficeOpenXml.ExcelPackage())
            {
                var objExcelFileRow = new clsExportedExcelFileHeaderRow();

                // Define the exported Excel file header row labels:
                objExcelFileRow.ESITPMID = "ID";
                objExcelFileRow.Customer = "Customer";
                objExcelFileRow.Course = "Course";
                objExcelFileRow.Modality = "Modality";
                objExcelFileRow.NumberStudents = "No. Students";
                objExcelFileRow.DeliveryCity = "Delivery City";
                objExcelFileRow.Language = "Language";
                objExcelFileRow.Location = "Location";
                objExcelFileRow.DeliveryWindowStart = "Delivery Window Start";
                objExcelFileRow.DeliveryWindowEnd = "Delivery Window End";
                objExcelFileRow.PreferedDateNotes = "Prefered Date Notes";
                objExcelFileRow.DeliveryChannel = "Delivery Channel";
                objExcelFileRow.ReadyToSchedule = "Ready To Schedule";
                
                // Create the Excel object in memory, name it,  and add a tab:
                var worksheet = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == "List");
                package.Workbook.Properties.Title = "Training Order List";
                worksheet = package.Workbook.Worksheets.Add("Training Order List");
                
                // Some formatting:
                worksheet.DefaultRowHeight = 12;
                worksheet.Row(1).Height = 20;

                // Populate header of Excel object:
                var intColumnCounter = 0;
                foreach (var v in objExcelFileRow.GetType().GetProperties())
                {
                    intColumnCounter ++ ;
                    worksheet.Cells[1, intColumnCounter].Value = v.GetValue(objExcelFileRow,null);
                    
                    worksheet.Column(intColumnCounter).AutoFit();
                }

                // Create the content type for the response stream:
                this.Response.Clear();
                this.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                this.Response.AddHeader(
                          "content-disposition",
                          string.Format("attachment;  filename={0}", "Training Order List.xlsx"));
                this.Response.BinaryWrite(package.GetAsByteArray());
                this.Response.End();

                // Back to home:
                return View("Index");
            }
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}