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

                var worksheetC2 = xlPackage.Workbook.Worksheets.Add("c2");
                workload.PrintToExcelC(worksheetC2, false);

                var worksheetO1 = xlPackage.Workbook.Worksheets.Add("o1");
                workload.PrintToExcelO(worksheetO1, false);

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
        public IActionResult BatchUserUpload(IFormFile batchUsers, string Faculty, TimeNorms timeNorms, string rowCount, int specialtiesRow, int RateRow)
        {
            workload.faculty = Faculty;
            workload.timeNorms = timeNorms;

            if (ModelState.IsValid)
            {
                if (batchUsers?.Length > 0)
                {
                    var stream = batchUsers.OpenReadStream();
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
                                    Console.WriteLine("Something went wrong"); // Вылавливание ошибки и вывод строки, в которой проихошла ошибка
                                }
                            }

                            for (var row = specialtiesRow + 1; row <= specialtiesRow + 4; row++)
                            {
                                string[] line = worksheet.Cells[row, 1].Value.ToString().Split(":", StringSplitOptions.RemoveEmptyEntries);
                                Dictionary<string, string> specializations = new Dictionary<string, string>();
                                foreach (string specialization in line[1].Split(",", StringSplitOptions.RemoveEmptyEntries))
                                {
                                    string[] NameAndCode = specialization.Split(" ", StringSplitOptions.RemoveEmptyEntries);
                                    specializations.Add(NameAndCode[0], NameAndCode[1]);
                                }
                                switch (row % specialtiesRow) {
                                    case 1:
                                        workload.Bachelor = specializations;
                                        break;
                                    case 2:
                                        workload.Specialty = specializations;
                                        break;
                                    case 3:
                                        workload.Magistracy = specializations;
                                        break;
                                    case 4:
                                        workload.Postgraduate = specializations;
                                        break;
                                }
                            }

                            // Ставки преподавателей
                            for (var row = RateRow + 1; worksheet.Cells[row, 1].Value != null; row++) {
                                workload.teacherRate.Add(worksheet.Cells[row, 1].Value.ToString(), 
                                    new Teacher(double.Parse(worksheet.Cells[row, 2].Value.ToString())));
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
