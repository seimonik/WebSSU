using Microsoft.AspNetCore.SignalR;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace WebSSU.Models
{
    public class TeachersWorkload
    {
        public class SubjectList
        {
            List<Subject> _subjects;
            public SubjectList()
            {
                _subjects = new List<Subject>();
            }
            public void Add(Subject subject)
            {
                _subjects.Add(subject);
            }
            public List<Subject> Get()
            {
                return _subjects;
            }
        }
        private Dictionary<string, SubjectList> teacherSubj = new Dictionary<string, SubjectList>();
        private Dictionary<string, SubjectList> theme = new Dictionary<string, SubjectList>();
        public TimeNorms timeNorms { get; set; }
        public string faculty { get; set; }
        // бакалавриат
        public Dictionary<string, string> Bachelor { get; set; }
        // специалитет
        public Dictionary<string, string> Specialty { get; set; }
        // магистратура
        public Dictionary<string, string> Magistracy { get; set; }
        // аспиранутра
        public Dictionary<string, string> Postgraduate { get; set; }

        // ставка преаодавателя 
        public Dictionary<string, Teacher> teacherRate = new Dictionary<string, Teacher>();


        private string LastSubjectName = " ";
        public void Add(string nameTeachers, Subject subject)
        {
            if (teacherSubj.ContainsKey(nameTeachers))
            {
                teacherSubj[nameTeachers].Add(subject);
            }
            else
            {
                teacherSubj.Add(nameTeachers, new SubjectList());
                teacherSubj[nameTeachers].Add(subject);
            }

            if (subject.Name == "--//--")
            {
                theme[LastSubjectName].Add(subject);
            }
            else
            {
                if (theme.ContainsKey(subject.Name))
                {
                    theme[subject.Name].Add(subject);
                    LastSubjectName = subject.Name;
                }
                else
                {
                    theme.Add(subject.Name, new SubjectList());
                    theme[subject.Name].Add(subject);
                    LastSubjectName = subject.Name;
                }
            }
        }
        public void PrintTableHeader(ExcelWorksheet worksheet, bool selfStudy, int startRow)
        {
            // Шрифт шапки таблицы
            worksheet.Cells[$"A{startRow}:Z{startRow + 2}"].Style.Font.Name = "Garamond";
            worksheet.Cells[$"A{startRow}:Z{startRow + 2}"].Style.Font.Size = 7;

            // Выравнивание по центру
            worksheet.Cells[$"A{startRow}:Z{startRow + 2}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[$"A{startRow}:Z{startRow + 2}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            // Перенос текста
            worksheet.Cells[$"A{startRow}:Z{startRow + 2}"].Style.WrapText = true;

            // ширина столбцов и высота строк заголовка
            worksheet.Row(9).Height = 51.8;
            worksheet.Column(1).Width = 28.22;
            worksheet.Column(2).Width = 17.78;
            for (int i = 3; i < 9; i++)
            {
                worksheet.Column(i).Width = 5.11;
            }
            for (int i = 9; i < 27; i++)
            {
                worksheet.Column(i).Width = 7.67;
            }

            worksheet.Cells[$"A{startRow}:A{startRow + 2}"].Merge = true;
            worksheet.Cells[$"A{startRow}"].Value = "Наименование дисциплины";

            // Поворот текста на 90
            worksheet.Cells[$"I{startRow + 1}:Y{startRow + 2}"].Style.TextRotation = 90;
            worksheet.Cells[$"B{startRow}:H{startRow + 2}"].Style.TextRotation = 90;

            // Слияние ячеек
            for (char c = 'B'; c <= 'H'; c++)
            {
                worksheet.Cells[$"{c}{startRow}:{c}{startRow + 2}"].Merge = true;
            }
            worksheet.Cells[$"B{startRow}"].Value = "Специальность или направление(код специальности или направления)";
            worksheet.Cells[$"C{startRow}"].Value = "Курс";
            worksheet.Cells[$"D{startRow}"].Value = "Семестр";
            worksheet.Cells[$"E{startRow}"].Value = "Количество студентов";
            worksheet.Cells[$"F{startRow}"].Value = "Количество потоков";
            worksheet.Cells[$"G{startRow}"].Value = "Количество групп/подгрупп";

            if (selfStudy)
            {
                worksheet.Cells[$"H{startRow}"].Value = "Самостоятельная работа по дисциплине в семестре (в часах)*";
            }

            worksheet.Cells[$"I{startRow}"].Value = "Число часов по видам учебной работы";
            worksheet.Cells[$"I{startRow}:Y{startRow}"].Merge = true;
            for (char c = 'I'; c <= 'Y'; c++)
            {
                worksheet.Cells[$"{c}{startRow + 1}:{c}{startRow + 2}"].Merge = true;
            }
            worksheet.Cells[$"I{startRow + 1}"].Value = "Лекции";
            worksheet.Cells[$"J{startRow + 1}"].Value = "Практ., семин. занятия";
            worksheet.Cells[$"K{startRow + 1}"].Value = "Лабор. занятия";
            worksheet.Cells[$"L{startRow + 1}"].Value = "Консультации по дисциплине, КСР";
            worksheet.Cells[$"M{startRow + 1}"].Value = "Консультации перед экзаменом";
            worksheet.Cells[$"N{startRow + 1}"].Value = "Экзамены";
            worksheet.Cells[$"O{startRow + 1}"].Value = "Зачеты";
            worksheet.Cells[$"P{startRow + 1}"].Value = "Руководство практикой";
            worksheet.Cells[$"Q{startRow + 1}"].Value = "Курсовые работы";
            worksheet.Cells[$"R{startRow + 1}"].Value = "Выпускные квалиф. работы";
            worksheet.Cells[$"S{startRow + 1}"].Value = "Работа в ГАК";
            worksheet.Cells[$"T{startRow + 1}"].Value = "Проверка контр. работ";
            worksheet.Cells[$"U{startRow + 1}"].Value = "Руководство аспирантами";
            worksheet.Cells[$"V{startRow + 1}"].Value = "Руководство соискателями";
            worksheet.Cells[$"W{startRow + 1}"].Value = "Руководство магис-терской программой";
            worksheet.Cells[$"X{startRow + 1}"].Value = "Факультативные занятия";

            worksheet.Cells[$"Z{startRow}:Z{startRow + 2}"].Merge = true;
            worksheet.Column(26).Width = 7.67;
            worksheet.Cells[$"Z{startRow}"].Value = "Итого (часов)";
            int col = 1;
            for (char c = 'A'; c <= 'Z'; c++)
            {
                worksheet.Cells[$"{c}{startRow + 3}"].Value = col;
                col++;
            }
            worksheet.Cells[$"A{startRow + 3}:Z{startRow + 3}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[$"A{startRow + 3}:Z{startRow + 3}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            var FirstTableRange = worksheet.Cells[$"A{startRow}:Z{startRow + 3}"];
            FirstTableRange.Style.Border.Top.Style = ExcelBorderStyle.Medium;
            FirstTableRange.Style.Border.Left.Style = ExcelBorderStyle.Medium;
            FirstTableRange.Style.Border.Right.Style = ExcelBorderStyle.Medium;
            FirstTableRange.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;

            worksheet.Cells[$"A{startRow + 4}:Z{startRow + 4}"].Merge = true;
            worksheet.Cells[$"A{startRow + 4}:Z{startRow + 4}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[$"A{startRow + 4}:Z{startRow + 4}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[$"A{startRow + 4}"].Value = "Очная форма обучения";
            worksheet.Cells[$"A{startRow + 4}"].Value = faculty;
        }
        public void PrintToExcelP(ExcelWorksheet worksheet, bool budget)
        {
            int row = 13;
            int startRow = row;
            TotalHours totalHours = new TotalHours();
            foreach(KeyValuePair<string, SubjectList> sub in theme)
            {
                foreach (Subject subject in sub.Value.Get())
                {
                    int StudentNumber;
                    if (budget)
                    {
                        StudentNumber = subject.Budget ?? 0;
                    }
                    else
                    {
                        StudentNumber = subject.Commercial ?? 0;
                        //if (StudentNumber == 0)
                        //    break;
                    }
                    if (budget || (!budget && StudentNumber > 0))
                    {
                        if (subject.Specialization == null || subject.Semester == null || subject.Budget == null
                        || subject.Commercial == null || subject.Groups == null)
                        {
                            worksheet.Cells[$"Z{row}"].Value = subject.TotalHours;
                        }

                        double Total = 0;

                        worksheet.Cells[$"A{row}"].Value = sub.Key;
                        worksheet.Cells[$"B{row}"].Value = subject.Specialization;
                        worksheet.Cells[$"C{row}"].Value = (subject.Semester % 2) + 1;
                        worksheet.Cells[$"D{row}"].Value = subject.Semester;
                        worksheet.Cells[$"E{row}"].Value = StudentNumber;
                        // количество потоков
                        int countFlow = subject.Specialization.Split(",", StringSplitOptions.RemoveEmptyEntries).Length;
                        worksheet.Cells[$"F{row}"].Value = countFlow;
                        // количество групп
                        int countGroups = subject.Groups.Split(",", StringSplitOptions.RemoveEmptyEntries).Length;
                        worksheet.Cells[$"G{row}"].Value = countGroups;
                        worksheet.Cells[$"H{row}"].Value = subject.SelfStudy;

                        // РУКОВОДСТВО ПРАКТИКОЙ
                        string nameLower = sub.Key.ToLower();
                        if (nameLower.Contains("практика"))
                        {
                            if (nameLower.Contains("базовая") || nameLower.Contains("технологическая") || nameLower.Contains("производственная"))
                            {
                                worksheet.Cells[$"P{row}"].Value = timeNorms.ProductionPractice * StudentNumber;
                                worksheet.Cells[$"Z{row}"].Value = timeNorms.ProductionPractice * StudentNumber;
                                totalHours.Practice += timeNorms.ProductionPractice * StudentNumber;
                                totalHours.Total += timeNorms.ProductionPractice * StudentNumber;
                            }
                            else if (nameLower.Contains("учебная") || nameLower.Contains("НИР"))
                            {
                                worksheet.Cells[$"P{row}"].Value = timeNorms.EducationalPractice * StudentNumber;
                                worksheet.Cells[$"Z{row}"].Value = timeNorms.EducationalPractice * StudentNumber;
                                totalHours.Practice += timeNorms.EducationalPractice * StudentNumber;
                                totalHours.Total += timeNorms.EducationalPractice * StudentNumber;
                            }
                            //else if (nameLower.Contains("технологическая") || nameLower.Contains("вычислительная"))
                            //{
                            //    worksheet.Cells[$"P{row}"].Value = 48 * countGroups;
                            //    worksheet.Cells[$"Z{row}"].Value = 48 * countGroups;
                            //    totalHours.Practice += 48 * countGroups;
                            //    totalHours.Total += 48 * countGroups;
                            //}
                            else if (nameLower.Contains("педагогическая") || nameLower.Contains("научно"))
                            {
                                worksheet.Cells[$"P{row}"].Value = 8 * StudentNumber;
                                worksheet.Cells[$"Z{row}"].Value = 8 * StudentNumber;
                                totalHours.Practice += 8 * StudentNumber;
                                totalHours.Total += 8 * StudentNumber;
                            }
                        }
                        else if (nameLower.Contains("консультация"))
                        {
                            worksheet.Cells[$"P{row}"].Value = 10 * StudentNumber;
                            worksheet.Cells[$"Z{row}"].Value = 10 * StudentNumber;
                            totalHours.Practice += 10 * StudentNumber;
                            totalHours.Total += 10 * StudentNumber;
                        }
                        // Работы в ГАК
                        else if (sub.Key.Contains("ГЭК"))
                        {
                            if (subject.ReportingForm.Contains("секретарь"))
                            {
                                double res = Math.Round((double)(2 * StudentNumber / 3), 1, MidpointRounding.AwayFromZero);
                                worksheet.Cells[$"S{row}"].Value = res;
                                worksheet.Cells[$"Z{row}"].Value = res;
                                totalHours.GAK += res;
                                totalHours.Total += res;
                            }
                            else
                            {
                                double res = Math.Round((double)(StudentNumber / 2), 1, MidpointRounding.AwayFromZero);
                                worksheet.Cells[$"S{row}"].Value = res;
                                worksheet.Cells[$"Z{row}"].Value = res;
                                totalHours.GAK += res;
                                totalHours.Total += res;
                            }
                        }

                        else if (sub.Key.Contains("ВКР"))
                        {
                            // ВКР
                            string[] specialization = subject.Specialization.Split(",", StringSplitOptions.RemoveEmptyEntries);
                            string[] groups = subject.Groups.Split(",", StringSplitOptions.RemoveEmptyEntries);
                            if (Magistracy.ContainsKey(specialization[0]))
                            {
                                // магистратура
                                worksheet.Cells[$"R{row}"].Value = timeNorms.FinalQualifyingWorkMagistracy * StudentNumber;
                                worksheet.Cells[$"Z{row}"].Value = timeNorms.FinalQualifyingWorkMagistracy * StudentNumber;
                                totalHours.VKR += timeNorms.FinalQualifyingWorkMagistracy * StudentNumber;
                                totalHours.Total += timeNorms.FinalQualifyingWorkMagistracy * StudentNumber;
                            }
                            else if (Bachelor.ContainsKey(specialization[0]))
                            {
                                // бакалавриат
                                worksheet.Cells[$"R{row}"].Value = timeNorms.FinalQualifyingWorkBachelor * StudentNumber;
                                worksheet.Cells[$"Z{row}"].Value = timeNorms.FinalQualifyingWorkBachelor * StudentNumber;
                                totalHours.VKR += timeNorms.FinalQualifyingWorkBachelor * StudentNumber;
                                totalHours.Total += timeNorms.FinalQualifyingWorkBachelor * StudentNumber;
                            }
                        }
                        else if (sub.Key.ToLower().Contains("курсовая работа"))
                        {
                            int year = ((subject.Semester ?? 0) % 2) + 1;
                            if (Bachelor.ContainsKey(subject.Specialization) && Bachelor[subject.Specialization] == "ПО" && year == 2)
                            {
                                worksheet.Cells[$"Q{row}"].Value = timeNorms.Coursework2PO * StudentNumber;
                                worksheet.Cells[$"Z{row}"].Value = timeNorms.Coursework2PO * StudentNumber;
                                totalHours.Coursework += timeNorms.Coursework2PO * StudentNumber;
                                totalHours.Total += timeNorms.Coursework2PO * StudentNumber;
                            }
                            else if (Bachelor.ContainsKey(subject.Specialization) && Bachelor[subject.Specialization] == "ПО" && year == 3)
                            {
                                worksheet.Cells[$"Q{row}"].Value = timeNorms.Coursework3PO * StudentNumber;
                                worksheet.Cells[$"Z{row}"].Value = timeNorms.Coursework3PO * StudentNumber;
                                totalHours.Coursework += timeNorms.Coursework3PO * StudentNumber;
                                totalHours.Total += timeNorms.Coursework3PO * StudentNumber;
                            }
                            else if (year == 2)
                            {
                                worksheet.Cells[$"Q{row}"].Value = timeNorms.Coursework2 * StudentNumber;
                                worksheet.Cells[$"Z{row}"].Value = timeNorms.Coursework2 * StudentNumber;
                                totalHours.Coursework += timeNorms.Coursework2 * StudentNumber;
                                totalHours.Total += timeNorms.Coursework2 * StudentNumber;
                            }
                            else if (year == 3)
                            {
                                worksheet.Cells[$"Q{row}"].Value = timeNorms.Coursework3 * StudentNumber;
                                worksheet.Cells[$"Z{row}"].Value = timeNorms.Coursework3 * StudentNumber;
                                totalHours.Coursework += timeNorms.Coursework3 * StudentNumber;
                                totalHours.Total += timeNorms.Coursework3 * StudentNumber;
                            }
                        }
                        else if (subject.Name == "--//--")
                        {
                            worksheet.Cells[$"J{row}"].Value = subject.Seminars;
                            Total += subject.Seminars == null ? 0 : subject.Seminars.Value;
                            worksheet.Cells[$"K{row}"].Value = subject.Laboratory;
                            Total += subject.Laboratory == null ? 0 : subject.Laboratory.Value;
                            worksheet.Cells[$"Z{row}"].Value = Total;

                            totalHours.Seminars += subject.Seminars ?? 0;
                            totalHours.Laboratory += subject.Laboratory ?? 0;
                            totalHours.Total += Total;

                            subject.Name = sub.Key;
                            subject.practice = true;
                        }
                        else if(budget || subject.GroupForm == "комм")
                        {
                            worksheet.Cells[$"I{row}"].Value = subject.Lectures;
                            Total += subject.Lectures == null ? 0 : subject.Lectures.Value;
                            totalHours.Lectures += subject.Lectures ?? 0;
                            worksheet.Cells[$"J{row}"].Value = subject.Seminars;
                            Total += subject.Seminars == null ? 0 : subject.Seminars.Value;
                            totalHours.Seminars += subject.Seminars ?? 0;
                            worksheet.Cells[$"K{row}"].Value = subject.Laboratory;
                            Total += subject.Laboratory == null ? 0 : subject.Laboratory.Value;
                            totalHours.Laboratory += subject.Laboratory ?? 0;
                            if (subject.SelfStudy != null)
                            {
                                double kons = Math.Round(((double)subject.SelfStudy * countFlow * 2.5 * 0.01), 1, MidpointRounding.AwayFromZero);
                                worksheet.Cells[$"L{row}"].Value = kons;
                                Total += kons;
                                totalHours.ConsultationsKSR += kons;
                            }

                            if (subject.ReportingForm == "зачет")
                            {
                                double zach = Math.Round(((double)StudentNumber / 3), 1, MidpointRounding.AwayFromZero);
                                worksheet.Cells[$"O{row}"].Value = zach;
                                Total += zach;
                                totalHours.Test += zach;
                            }
                            else if (subject.ReportingForm == "экзамен")
                            {
                                double ekz = Math.Round(((double)StudentNumber / 2), 1, MidpointRounding.AwayFromZero);
                                worksheet.Cells[$"M{row}"].Value = 2 * countFlow;
                                worksheet.Cells[$"N{row}"].Value = ekz;
                                Total += ekz + 2 * countFlow;
                                totalHours.Exams += ekz;
                                totalHours.Consultations += 2 * countFlow;
                            }

                            // Проверка контрольных работ
                            double kr = Math.Round(((double)StudentNumber / 6), 1, MidpointRounding.AwayFromZero);
                            worksheet.Cells[$"T{row}"].Value = kr;
                            Total += kr;
                            totalHours.ControlWorks += kr;

                            if (sub.Key.Trim() == "Руководство аспирантами")
                            {
                                worksheet.Cells[$"W{row}"].Value = subject.TotalHours;
                                Total += subject.TotalHours == null ? 0 : subject.TotalHours.Value;
                                totalHours.LeadershipGraduateStudents += subject.TotalHours ?? 0;
                            }

                            // Итого часов
                            worksheet.Cells[$"Z{row}"].Value = Total;
                            totalHours.Total += Total;
                        }
                        row++;
                    }
                }
                //row++;
            }
            worksheet.Cells[$"A{startRow}:A{row}"].Style.WrapText = true;
            worksheet.Cells[$"A{startRow}:A{row}"].Style.Font.Name = "Times New Roman";
            worksheet.Cells[$"A{startRow}:A{row}"].Style.Font.Size = 10;

            worksheet.Cells[$"A{row}"].Value = "Итого по " + faculty;
            totalHours.PrintToExcel(worksheet, row);
        }

        public void PrintC(ExcelWorksheet worksheet, bool budget, List<Subject> subjects, ref int row, string teacherName)
        {
            TotalHours totalHours = new TotalHours();
            foreach (Subject subject in subjects)
            {
                int StudentNumber;
                if (budget)
                {
                    StudentNumber = subject.Budget ?? 0;
                }
                else
                {
                    StudentNumber = subject.Commercial ?? 0;
                    //if (StudentNumber == 0)
                    //    break;
                }
                if (budget || (!budget && subject.Commercial != null && subject.Commercial > 0))
                {
                    if (subject.Specialization == null || subject.Semester == null || subject.Budget == null
                    || subject.Commercial == null || subject.Groups == null)
                    {
                        worksheet.Cells[$"A{row}"].Value = subject.Name;
                        worksheet.Cells[$"Z{row}"].Value = subject.TotalHours;
                    }
                    double Total = 0;

                    worksheet.Cells[$"A{row}"].Value = subject.Name;
                    worksheet.Cells[$"B{row}"].Value = subject.Specialization;
                    worksheet.Cells[$"C{row}"].Value = (subject.Semester % 2) + 1;
                    worksheet.Cells[$"D{row}"].Value = subject.Semester;
                    worksheet.Cells[$"E{row}"].Value = StudentNumber;
                    // количество потоков
                    int countFlow = subject.Specialization.Split(",", StringSplitOptions.RemoveEmptyEntries).Length;
                    worksheet.Cells[$"F{row}"].Value = countFlow;
                    // количество групп
                    int countGroups = subject.Groups.Split(",", StringSplitOptions.RemoveEmptyEntries).Length;
                    worksheet.Cells[$"G{row}"].Value = countGroups;
                    worksheet.Cells[$"H{row}"].Value = subject.SelfStudy;

                    // РУКОВОДСТВО ПРАКТИКОЙ
                    string nameLower = subject.Name.ToLower();
                    if (nameLower.Contains("практика"))
                    {
                        if (nameLower.Contains("базовая") || nameLower.Contains("производственная"))
                        {
                            worksheet.Cells[$"P{row}"].Value = 2 * StudentNumber;
                            worksheet.Cells[$"Z{row}"].Value = 2 * StudentNumber;
                            totalHours.Practice += 2 * StudentNumber;
                            totalHours.Total += 2 * StudentNumber;
                        }
                        else if (nameLower.Contains("технологическая") || nameLower.Contains("вычислительная"))
                        {
                            worksheet.Cells[$"P{row}"].Value = 48 * countGroups;
                            worksheet.Cells[$"Z{row}"].Value = 48 * countGroups;
                            totalHours.Practice += 48 * countGroups;
                            totalHours.Total += 48 * countGroups;
                        }
                        else if (nameLower.Contains("педагогическая") || nameLower.Contains("научно"))
                        {
                            worksheet.Cells[$"P{row}"].Value = 8 * StudentNumber;
                            worksheet.Cells[$"Z{row}"].Value = 8 * StudentNumber;
                            totalHours.Practice += 8 * StudentNumber;
                            totalHours.Total += 8 * StudentNumber;
                        }
                    }
                    else if (nameLower.Contains("консультация"))
                    {
                        worksheet.Cells[$"P{row}"].Value = 10 * StudentNumber;
                        worksheet.Cells[$"Z{row}"].Value = 10 * StudentNumber;
                        totalHours.Practice += 10 * StudentNumber;
                        totalHours.Total += 10 * StudentNumber;
                    }


                    else if (subject.Name.Contains("ВКР"))
                    {
                        // ВКР
                        string[] groups = subject.Groups.Split(",", StringSplitOptions.RemoveEmptyEntries);
                        if ((int.Parse(groups[0]) / 10) % 10 == 7)
                        {
                            // магистратура
                            worksheet.Cells[$"R{row}"].Value = 34 * StudentNumber;
                            worksheet.Cells[$"Z{row}"].Value = 34 * StudentNumber;
                            totalHours.VKR += 34 * StudentNumber;
                            totalHours.Total += 34 * StudentNumber;
                        }
                        else if ((int.Parse(groups[0]) / 10) % 10 == 3)
                        {
                            // специалитет
                            worksheet.Cells[$"R{row}"].Value = 30 * StudentNumber;
                            worksheet.Cells[$"Z{row}"].Value = 30 * StudentNumber;
                            totalHours.VKR += 30 * StudentNumber;
                            totalHours.Total += 30 * StudentNumber;
                        }
                        else
                        {
                            worksheet.Cells[$"R{row}"].Value = 24 * StudentNumber;
                            worksheet.Cells[$"Z{row}"].Value = 24 * StudentNumber;
                            totalHours.VKR += 24 * StudentNumber;
                            totalHours.Total += 24 * StudentNumber;
                        }
                    }
                    else if (subject.Name.ToLower().Contains("курсовая работа"))
                    {
                        // КУРСОВЫЕ РАБОТЫ
                        string[] note = subject.Remark.Split(" ", StringSplitOptions.RemoveEmptyEntries);
                        worksheet.Cells[$"Q{row}"].Value = int.Parse(note[0]) * StudentNumber;
                        worksheet.Cells[$"Z{row}"].Value = int.Parse(note[0]) * StudentNumber;
                        totalHours.Coursework += int.Parse(note[0]) * StudentNumber;
                        totalHours.Total += int.Parse(note[0]) * StudentNumber;
                    }
                    else if (subject.practice)//(subject.Name == "--//--")
                    {
                        worksheet.Cells[$"J{row}"].Value = subject.Seminars;
                        Total += subject.Seminars == null ? 0 : subject.Seminars.Value;
                        worksheet.Cells[$"K{row}"].Value = subject.Laboratory;
                        Total += subject.Laboratory == null ? 0 : subject.Laboratory.Value;
                        worksheet.Cells[$"Z{row}"].Value = Total;

                        totalHours.Seminars += subject.Seminars ?? 0;
                        totalHours.Laboratory += subject.Laboratory ?? 0;
                        totalHours.Total += Total;
                    }
                    else
                    {
                        worksheet.Cells[$"I{row}"].Value = subject.Lectures;
                        Total += subject.Lectures == null ? 0 : subject.Lectures.Value;
                        totalHours.Lectures += subject.Lectures ?? 0;
                        worksheet.Cells[$"J{row}"].Value = subject.Seminars;
                        Total += subject.Seminars == null ? 0 : subject.Seminars.Value;
                        totalHours.Seminars += subject.Seminars ?? 0;
                        worksheet.Cells[$"K{row}"].Value = subject.Laboratory;
                        Total += subject.Laboratory == null ? 0 : subject.Laboratory.Value;
                        totalHours.Laboratory += subject.Laboratory ?? 0;
                        if (subject.SelfStudy != null)
                        {
                            double kons = Math.Round(((double)subject.SelfStudy * countFlow * 2.5 * 0.01), 1, MidpointRounding.AwayFromZero);
                            worksheet.Cells[$"L{row}"].Value = kons;
                            Total += kons;
                            totalHours.ConsultationsKSR += kons;
                        }

                        if (subject.ReportingForm == "зачет")
                        {
                            double zach = Math.Round(((double)StudentNumber / 3), 1, MidpointRounding.AwayFromZero);
                            worksheet.Cells[$"O{row}"].Value = zach;
                            Total += zach;
                            totalHours.Test += zach;
                        }
                        else if (subject.ReportingForm == "экзамен")
                        {
                            double ekz = Math.Round(((double)StudentNumber / 2), 1, MidpointRounding.AwayFromZero);
                            worksheet.Cells[$"M{row}"].Value = 2 * countFlow;
                            worksheet.Cells[$"N{row}"].Value = ekz;
                            Total += ekz + 2 * countFlow;
                            totalHours.Exams += ekz;
                            totalHours.Consultations += 2 * countFlow;
                        }

                        // Работы в ГАК
                        // ???????

                        // Проверка контрольных работ
                        double kr = Math.Round(((double)StudentNumber / 6), 1, MidpointRounding.AwayFromZero);
                        worksheet.Cells[$"T{row}"].Value = kr;
                        Total += kr;
                        totalHours.ControlWorks += kr;

                        if (subject.Name.Trim() == "Руководство аспирантами")
                        {
                            worksheet.Cells[$"W{row}"].Value = subject.TotalHours;
                            Total += subject.TotalHours == null ? 0 : subject.TotalHours.Value;
                            totalHours.LeadershipGraduateStudents += subject.TotalHours ?? 0;
                        }

                        // Итого часов
                        worksheet.Cells[$"Z{row}"].Value = Total;
                        totalHours.Total += Total;
                    }
                    row++;
                }
                worksheet.Cells[$"A{row}"].Value = "Итого по " + faculty;
                totalHours.PrintToExcel(worksheet, row);
                if (budget)
                {
                    teacherRate[teacherName].AmountHoursBudget = totalHours;
                }
                else
                {
                    teacherRate[teacherName].AmountHoursCommercial = totalHours;
                }
            }
        }

        public void PrintToExcelC(ExcelWorksheet worksheet, bool budget)
        {
            int row = 1;
            foreach (KeyValuePair<string, SubjectList> teacher in teacherSubj)
            {
                worksheet.Cells[$"A{row}"].Value = "Карточка учебных поручений на 2018/ 2019 учебный год";
                worksheet.Cells[$"A{row}:I{row}"].Merge = true;
                worksheet.Cells[$"A{++row}"].Value = "Фамилия, имя, отчество преподавателя " + teacher.Key;
                worksheet.Cells[$"A{row}:I{row}"].Merge = true;
                worksheet.Cells[$"A{++row}"].Value = "Ученая степень, ученое звание ____________________________________";
                worksheet.Cells[$"A{row}:I{row}"].Merge = true;
                worksheet.Cells[$"N{row}"].Value = "Форма нагрузки бюджетная";
                worksheet.Cells[$"N{row}:S{row}"].Merge = true;
                worksheet.Cells[$"A{++row}"].Value = "Должность, ставка";
                worksheet.Cells[$"A{row}:I{row}"].Merge = true;
                worksheet.Cells[$"N{row}"].Value = "Кафедра  информатики и программирования";
                worksheet.Cells[$"N{row}:S{row}"].Merge = true;
                worksheet.Cells[$"A{++row}"].Value = "Основная, внутреннее совмещение, внешнее совмещение, почасовая оплата";
                worksheet.Cells[$"A{row}:I{row}"].Merge = true;
                worksheet.Cells[$"N{row}"].Value = "Факультет  КНиИТ";
                worksheet.Cells[$"N{row}:S{row}"].Merge = true;
                worksheet.Cells[$"A{++row}"].Value = "(нужное подчеркнуть)";
                worksheet.Cells[$"A{row}:I{row}"].Merge = true;

                // 1 семестр
                worksheet.Cells[$"A{++row}"].Value = "1 семестр";
                worksheet.Cells[$"A{row}:Z{row}"].Merge = true;
                PrintTableHeader(worksheet, false, ++row);

                row += 5;

                worksheet.Cells[$"A{++row}"].Value = faculty;

                int startRow = row;

                var subjects = teacher.Value.Get().Where(s => s.Semester % 2 == 1).ToList();
                PrintC(worksheet, budget, subjects, ref row, teacher.Key);

                worksheet.Cells[$"A{startRow}:A{row}"].Style.WrapText = true;
                worksheet.Cells[$"A{startRow}:A{row}"].Style.Font.Name = "Times New Roman";
                worksheet.Cells[$"A{startRow}:A{row}"].Style.Font.Size = 10;

                // 2 семестр
                worksheet.Cells[$"A{++row}"].Value = "2 семестр";
                worksheet.Cells[$"A{row}:Z{row}"].Merge = true;
                PrintTableHeader(worksheet, false, ++row);

                row += 5;

                worksheet.Cells[$"A{++row}"].Value = faculty;

                startRow = row;

                subjects = teacher.Value.Get().Where(s => s.Semester % 2 == 0).ToList();
                PrintC(worksheet, budget, subjects, ref row, teacher.Key);

                worksheet.Cells[$"A{startRow}:A{row}"].Style.WrapText = true;
                worksheet.Cells[$"A{startRow}:A{row}"].Style.Font.Name = "Times New Roman";
                worksheet.Cells[$"A{startRow}:A{row}"].Style.Font.Size = 10;

                row += 4;
            }
        }

        public void PrintTableHeaderO(ExcelWorksheet worksheet, int startRow)
        {
            // Шрифт шапки таблицы
            worksheet.Cells[$"A{startRow}:Z{startRow + 2}"].Style.Font.Name = "Garamond";
            worksheet.Cells[$"A{startRow}:Z{startRow + 2}"].Style.Font.Size = 7;

            // Выравнивание по центру
            worksheet.Cells[$"A{startRow}:Z{startRow + 2}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[$"A{startRow}:Z{startRow + 2}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            // Перенос текста
            worksheet.Cells[$"A{startRow}:Z{startRow + 2}"].Style.WrapText = true;

            // ширина столбцов и высота строк заголовка
            worksheet.Row(9).Height = 51.8;
            worksheet.Column(1).Width = 28.22;
            worksheet.Column(2).Width = 17.78;
            for (int i = 3; i < 9; i++)
            {
                worksheet.Column(i).Width = 5.11;
            }
            for (int i = 9; i < 27; i++)
            {
                worksheet.Column(i).Width = 7.67;
            }

            worksheet.Cells[$"A{startRow}:A{startRow + 2}"].Merge = true;
            worksheet.Cells[$"A{startRow}"].Value = "Фамилия, инициалы преподавателя, его должность";

            // Поворот текста на 90
            worksheet.Cells[$"C{startRow + 1}:S{startRow + 2}"].Style.TextRotation = 90;
            worksheet.Cells[$"B{startRow}"].Style.TextRotation = 90;

            // Слияние ячеек
            worksheet.Cells[$"B{startRow}:B{startRow + 2}"].Merge = true;
            
            worksheet.Cells[$"B{startRow}"].Value = "Форма обучения (очная, очно-заочная, заочная)";

            worksheet.Cells[$"C{startRow}"].Value = "Число часов по видам учебной работы";
            worksheet.Cells[$"C{startRow}:S{startRow}"].Merge = true;
            for (char c = 'C'; c <= 'Y'; c++)
            {
                worksheet.Cells[$"{c}{startRow + 1}:{c}{startRow + 2}"].Merge = true;
            }
            worksheet.Cells[$"C{startRow + 1}"].Value = "Лекции";
            worksheet.Cells[$"D{startRow + 1}"].Value = "Практ., семин. занятия";
            worksheet.Cells[$"E{startRow + 1}"].Value = "Лабор. занятия";
            worksheet.Cells[$"F{startRow + 1}"].Value = "Консультации по дисциплине, КСР";
            worksheet.Cells[$"G{startRow + 1}"].Value = "Консультации перед экзаменом";
            worksheet.Cells[$"H{startRow + 1}"].Value = "Экзамены";
            worksheet.Cells[$"I{startRow + 1}"].Value = "Зачеты";
            worksheet.Cells[$"J{startRow + 1}"].Value = "Руководство практикой";
            worksheet.Cells[$"K{startRow + 1}"].Value = "Курсовые работы";
            worksheet.Cells[$"L{startRow + 1}"].Value = "Выпускные квалиф. работы";
            worksheet.Cells[$"M{startRow + 1}"].Value = "Работа в ГАК";
            worksheet.Cells[$"N{startRow + 1}"].Value = "Проверка контр. работ";
            worksheet.Cells[$"O{startRow + 1}"].Value = "Руководство аспирантами";
            worksheet.Cells[$"P{startRow + 1}"].Value = "Руководство соискателями";
            worksheet.Cells[$"Q{startRow + 1}"].Value = "Руководство магис-терской программой";
            worksheet.Cells[$"R{startRow + 1}"].Value = "Факультативные занятия";

            worksheet.Cells[$"T{startRow}:T{startRow + 2}"].Merge = true;
            worksheet.Column(26).Width = 7.67;
            worksheet.Cells[$"T{startRow}"].Value = "Итого (часов)";
            int col = 1;
            for (char c = 'A'; c <= 'T'; c++)
            {
                worksheet.Cells[$"{c}{startRow + 3}"].Value = col;
                col++;
            }
            worksheet.Cells[$"A{startRow + 3}:T{startRow + 3}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[$"A{startRow + 3}:T{startRow + 3}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            var FirstTableRange = worksheet.Cells[$"A{startRow}:T{startRow + 3}"];
            FirstTableRange.Style.Border.Top.Style = ExcelBorderStyle.Medium;
            FirstTableRange.Style.Border.Left.Style = ExcelBorderStyle.Medium;
            FirstTableRange.Style.Border.Right.Style = ExcelBorderStyle.Medium;
            FirstTableRange.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;

            worksheet.Cells[$"A{startRow + 4}:T{startRow + 4}"].Merge = true;
            worksheet.Cells[$"A{startRow + 4}:T{startRow + 4}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[$"A{startRow + 4}:T{startRow + 4}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        }

        public void PrintToExcelO(ExcelWorksheet worksheet, bool budget)
        {
            int row = 8;
            PrintTableHeaderO(worksheet, row);
            row += 4;
            foreach(KeyValuePair<string, Teacher> teacher in teacherRate)
            {
                TotalHours total = teacher.Value.AmountHoursBudget;
                worksheet.Cells[$"A{row}"].Value = teacher.Key;
                worksheet.Cells[$"B{row}"].Value = "очн";
                worksheet.Cells[$"C{row}"].Value = total.Lectures;
                worksheet.Cells[$"D{row}"].Value = total.Seminars;
                worksheet.Cells[$"E{row}"].Value = total.Laboratory;
                worksheet.Cells[$"F{row}"].Value = total.ConsultationsKSR;
                worksheet.Cells[$"G{row}"].Value = total.Consultations;
                worksheet.Cells[$"H{row}"].Value = total.Exams;
                worksheet.Cells[$"I{row}"].Value = total.Test;
                worksheet.Cells[$"J{row}"].Value = total.Practice;
                worksheet.Cells[$"K{row}"].Value = total.Coursework;
                worksheet.Cells[$"L{row}"].Value = total.VKR;
                worksheet.Cells[$"M{row}"].Value = total.GAK;
                worksheet.Cells[$"N{row}"].Value = total.ControlWorks;
                worksheet.Cells[$"O{row}"].Value = total.LeadershipGraduateStudents;
                worksheet.Cells[$"P{row}"].Value = total.LeadershipApplicants;
                worksheet.Cells[$"Q{row}"].Value = total.LeadershipMaster;
                worksheet.Cells[$"R{row}"].Value = total.OptionalClasses;
                worksheet.Cells[$"T{row}"].Value = total.Total;

                row++;
            }
        }
    }
}
