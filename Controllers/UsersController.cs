using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using System.Drawing;
using WebSSU.Models;
using System.Text.RegularExpressions;

namespace WebSSU.Controllers
{
    public class UsersController : Controller
    {
        private readonly ILogger<UsersController> _logger;
        TeachersWorkload workload = new TeachersWorkload();
        //List<Teacher> teachers = new List<Teacher>();

        public UsersController(ILogger<UsersController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            var users = GetlistOfUsers();

            return View(users);
        }

        public IActionResult ExportToExcel()
        {
            // Get the user list 
            var users = GetlistOfUsers();

            var stream = new MemoryStream();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var xlPackage = new ExcelPackage(stream))
            {
                var worksheetP1 = xlPackage.Workbook.Worksheets.Add("p1");
                var namedStyle = xlPackage.Workbook.Styles.CreateNamedStyle("HyperLink");
                namedStyle.Style.Font.UnderLine = true;
                namedStyle.Style.Font.Color.SetColor(Color.Blue);
                const int startRow = 5;
                var row = startRow;

                workload.PrintTableHeader(worksheetP1, true, 7);
                workload.PrintToExcelP(worksheetP1, true);

                var worksheetP2 = xlPackage.Workbook.Worksheets.Add("p2");
                workload.PrintTableHeader(worksheetP2, true, 7);
                workload.PrintToExcelP(worksheetP2, false);

                var worksheetC1 = xlPackage.Workbook.Worksheets.Add("c1");
                workload.PrintToExcelC(worksheetC1, true);

                // set some core property values
                xlPackage.Workbook.Properties.Title = "User List";
                xlPackage.Workbook.Properties.Author = "WebSSU";
                xlPackage.Workbook.Properties.Subject = "User List";
                // save the new spreadsheet
                xlPackage.Save();
                // Response.Clear();
            }
            stream.Position = 0;
            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "users.xlsx");
        }

        [HttpGet]
        public IActionResult BatchUserUpload()
        {
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult BatchUserUpload(IFormFile batchUsers, string Faculty, string rowCount)
        {
            workload.faculty = Faculty;

            if (ModelState.IsValid)
            {
                if (batchUsers?.Length > 0)
                {
                    var stream = batchUsers.OpenReadStream();
                    List<User> users = new List<User>();
                    //List<Teacher> teachers = new List<Teacher>();
                    try
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        using (var package = new ExcelPackage(stream))
                        {
                            var worksheet = package.Workbook.Worksheets.First();//package.Workbook.Worksheets[0];
                            //var rowCount = worksheet.Dimension.Rows;
                            int rowTotal = int.Parse(rowCount);

                            for (var row = 9; row <= rowTotal; row++) // <= rowCount
                            {
                                try
                                {
                                    // НЕСКОЛЬКО СЕМЕСТРОВ
                                    //List<int> semestrs = new List<int>();
                                    //if (worksheet.Cells[row, 3].Value != null)
                                    //{
                                    //    foreach (string sem in worksheet.Cells[row, 3].Value.ToString().Split(",", StringSplitOptions.RemoveEmptyEntries))
                                    //    {
                                    //        semestrs.Add(int.Parse(sem));
                                    //    }
                                    //}
                                    string nameTeacher = worksheet.Cells[row, 16].Value?.ToString() == null ? " " : worksheet.Cells[row, 16].Value.ToString();
                                    Subject subject = new Subject()
                                    {
                                        Name = worksheet.Cells[row, 1].Value?.ToString(),
                                        Specialization = worksheet.Cells[row, 2].Value.ToString(),
                                        Semester = int.Parse(worksheet.Cells[row, 3].Value.ToString()),
                                        Budget = worksheet.Cells[row, 4].Value == null ? 0 : int.Parse(worksheet.Cells[row, 4].Value.ToString()),
                                        Commercial = worksheet.Cells[row, 5].Value == null ? 0 : int.Parse(worksheet.Cells[row, 5].Value.ToString()),
                                        Groups = worksheet.Cells[row, 6].Value.ToString(),
                                        GroupForm = worksheet.Cells[row, 7].Value.ToString(),
                                        TotalHours = worksheet.Cells[row, 8].Value == null ? null : int.Parse(worksheet.Cells[row, 8].Value?.ToString()),
                                        Lectures = worksheet.Cells[row, 9].Value == null ? null : int.Parse(worksheet.Cells[row, 9].Value?.ToString()),
                                        Seminars = worksheet.Cells[row, 10].Value == null ? null : int.Parse(worksheet.Cells[row, 10].Value?.ToString()),
                                        Laboratory = worksheet.Cells[row, 11].Value == null ? null : int.Parse(worksheet.Cells[row, 11].Value?.ToString()),
                                        SelfStudy = worksheet.Cells[row, 12].Value == null ? null : int.Parse(worksheet.Cells[row, 12].Value?.ToString()),
                                        LoadPerWeek = worksheet.Cells[row, 13].Value == null ? null : int.Parse(worksheet.Cells[row, 13].Value?.ToString()),
                                        ReportingForm = worksheet.Cells[row, 14].Value?.ToString(),
                                        Remark = worksheet.Cells[row, 15].Value?.ToString()
                                    };
                                    workload.Add(nameTeacher, subject);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("Something went wrong");
                                }
                            }
                        }
                        return ExportToExcel();
                        //return View("Index", users);

                    }
                    catch (Exception e)
                    {
                        return View();
                    }
                }
            }

            return View();
        }

        // Mimic a database operation
        private List<User> GetlistOfUsers()
        {
            var users = new List<User>()
        {
            new User {
                Email = "mohamad@email.com",
                Name = "Mohamad",
                Phone = "123456"
            },
            new User {
                Email = "donald@email.com",
                Name = "donald",
                Phone = "222222"
            },
            new User {
                Email = "mickey@email.com",
                Name = "mickey",
                Phone = "33333"
            }
        };

            return users;
        }
    }
}
