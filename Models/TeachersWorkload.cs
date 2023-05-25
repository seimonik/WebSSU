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
        public string faculty { get; set; }

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
            worksheet.Cells[$"B{startRow + 1}:Y{startRow + 2}"].Style.TextRotation = 90;
            worksheet.Cells[$"B{startRow + 1}:H{startRow + 2}"].Style.TextRotation = 90;

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
                    if (budget || (!budget && subject.Commercial != null && subject.Commercial > 0))
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


                        else if (sub.Key.Contains("ВКР"))
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
                        else if (sub.Key.ToLower().Contains("курсовая работа"))
                        {
                            // КУРСОВЫЕ РАБОТЫ
                            string[] note = subject.Remark.Split(" ", StringSplitOptions.RemoveEmptyEntries);
                            worksheet.Cells[$"Q{row}"].Value = int.Parse(note[0]) * StudentNumber;
                            worksheet.Cells[$"Z{row}"].Value = int.Parse(note[0]) * StudentNumber;
                            totalHours.Coursework += int.Parse(note[0]) * StudentNumber;
                            totalHours.Total += int.Parse(note[0]) * StudentNumber;
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
            worksheet.Cells[$"A{row}"].Value = "Итого по " + faculty;
            totalHours.PrintToExcel(worksheet, row);
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

                worksheet.Cells[$"A{++row}"].Value = "1 семестр";
                worksheet.Cells[$"A{row}:Z{row}"].Merge = true;
                PrintTableHeader(worksheet, false, ++row);

                worksheet.Cells[$"A{++row}"].Value = faculty;

                TotalHours totalHours = new TotalHours();
                foreach (Subject subject in teacher.Value.Get())
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
                }

                row += 8;
                row++;
            }
        }

        //public void PrintToExcelP2(ExcelWorksheet worksheet)
        //{
        //    int row = 13;
        //    TotalHours totalHours = new TotalHours();
        //    foreach (KeyValuePair<string, SubjectList> sub in theme)
        //    {
        //        foreach (Subject subject in sub.Value.Get())
        //        {
        //            if (subject.Commercial != null && subject.Commercial > 0)
        //            {
        //                //if (subject.Specialization == null || subject.Semester == null 
        //                //    || subject.Commercial == null || subject.Groups == null)
        //                //{
        //                //    worksheet.Cells[$"A{row}"].Value = sub.Key;
        //                //    worksheet.Cells[$"Z{row}"].Value = subject.TotalHours;
        //                //}

        //                double Total = 0;

        //                worksheet.Cells[$"A{row}"].Value = sub.Key;
        //                worksheet.Cells[$"B{row}"].Value = subject.Specialization;
        //                worksheet.Cells[$"C{row}"].Value = (subject.Semester % 2) + 1;
        //                worksheet.Cells[$"D{row}"].Value = subject.Semester;
        //                worksheet.Cells[$"E{row}"].Value = subject.Commercial;
        //                // количество потоков
        //                int countFlow = subject.Specialization.Split(",", StringSplitOptions.RemoveEmptyEntries).Length;
        //                worksheet.Cells[$"F{row}"].Value = countFlow;
        //                // количество групп
        //                int countGroups = subject.Groups.Split(",", StringSplitOptions.RemoveEmptyEntries).Length;
        //                worksheet.Cells[$"G{row}"].Value = countGroups;
        //                worksheet.Cells[$"H{row}"].Value = subject.SelfStudy;

        //                // РУКОВОДСТВО ПРАКТИКОЙ
        //                string nameLower = sub.Key.ToLower();
        //                if (nameLower.Contains("практика"))
        //                {
        //                    if (nameLower.Contains("базовая") || nameLower.Contains("производственная"))
        //                    {
        //                        worksheet.Cells[$"P{row}"].Value = 2 * subject.Commercial;
        //                        worksheet.Cells[$"Z{row}"].Value = 2 * subject.Commercial;
        //                        totalHours.Practice += 2 * subject.Commercial ?? 0;
        //                        totalHours.Total += 2 * subject.Commercial ?? 0;
        //                    }
        //                    else if (nameLower.Contains("технологическая") || nameLower.Contains("вычислительная"))
        //                    {
        //                        worksheet.Cells[$"P{row}"].Value = 48 * countGroups;
        //                        worksheet.Cells[$"Z{row}"].Value = 48 * countGroups;
        //                        totalHours.Practice += 48 * countGroups;
        //                        totalHours.Total += 48 * countGroups;
        //                    }
        //                    else if (nameLower.Contains("педагогическая") || nameLower.Contains("научно"))
        //                    {
        //                        worksheet.Cells[$"P{row}"].Value = 8 * subject.Commercial;
        //                        worksheet.Cells[$"Z{row}"].Value = 8 * subject.Commercial;
        //                        totalHours.Practice += 8 * subject.Commercial ?? 0;
        //                        totalHours.Total += 8 * subject.Commercial ?? 0;
        //                    }
        //                }
        //                else if (nameLower.Contains("консультация"))
        //                {
        //                    worksheet.Cells[$"P{row}"].Value = 10 * subject.Commercial;
        //                    worksheet.Cells[$"Z{row}"].Value = 10 * subject.Commercial;
        //                    totalHours.Practice += 10 * subject.Commercial ?? 0;
        //                    totalHours.Total += 10 * subject.Commercial ?? 0;
        //                }


        //                else if (sub.Key.Contains("ВКР"))
        //                {
        //                    // ВКР
        //                    string[] groups = subject.Groups.Split(",", StringSplitOptions.RemoveEmptyEntries);
        //                    if ((int.Parse(groups[0]) / 10) % 10 == 7)
        //                    {
        //                        // магистратура
        //                        worksheet.Cells[$"R{row}"].Value = 34 * subject.Commercial;
        //                        worksheet.Cells[$"Z{row}"].Value = 34 * subject.Commercial  ;
        //                        totalHours.VKR += 34 * subject.Commercial ?? 0;
        //                        totalHours.Total += 34 * subject.Commercial ?? 0;
        //                    }
        //                    else if ((int.Parse(groups[0]) / 10) % 10 == 3)
        //                    {
        //                        // специалитет
        //                        worksheet.Cells[$"R{row}"].Value = 30 * subject.Commercial;
        //                        worksheet.Cells[$"Z{row}"].Value = 30 * subject.Commercial;
        //                        totalHours.VKR += 30 * subject.Commercial ?? 0;
        //                        totalHours.Total += 30 * subject.Commercial ?? 0;
        //                    }
        //                    else
        //                    {
        //                        worksheet.Cells[$"R{row}"].Value = 24 * subject.Commercial;
        //                        worksheet.Cells[$"Z{row}"].Value = 24 * subject.Commercial;
        //                        totalHours.VKR += 24 * subject.Commercial ?? 0;
        //                        totalHours.Total += 24 * subject.Commercial ?? 0;
        //                    }
        //                }
        //                else if (sub.Key.ToLower().Contains("курсовая работа"))
        //                {
        //                    // КУРСОВЫЕ РАБОТЫ
        //                    string[] note = subject.Remark.Split(" ", StringSplitOptions.RemoveEmptyEntries);
        //                    worksheet.Cells[$"Q{row}"].Value = int.Parse(note[0]) * subject.Commercial;
        //                    worksheet.Cells[$"Z{row}"].Value = int.Parse(note[0]) * subject.Commercial;
        //                    totalHours.Coursework += int.Parse(note[0]) * subject.Commercial ?? 0;
        //                    totalHours.Total += int.Parse(note[0]) * subject.Commercial ?? 0;
        //                }
        //                else if (subject.Name == "--//--")
        //                {
        //                    worksheet.Cells[$"J{row}"].Value = subject.Seminars;
        //                    Total += subject.Seminars == null ? 0 : subject.Seminars.Value;
        //                    worksheet.Cells[$"K{row}"].Value = subject.Laboratory;
        //                    Total += subject.Laboratory == null ? 0 : subject.Laboratory.Value;
        //                    worksheet.Cells[$"Z{row}"].Value = Total;

        //                    totalHours.Seminars += subject.Seminars ?? 0;
        //                    totalHours.Laboratory += subject.Laboratory ?? 0;
        //                    totalHours.Total += Total;
        //                }
        //                else
        //                {
        //                    worksheet.Cells[$"I{row}"].Value = subject.Lectures;
        //                    Total += subject.Lectures == null ? 0 : subject.Lectures.Value;
        //                    totalHours.Lectures += subject.Lectures ?? 0;
        //                    worksheet.Cells[$"J{row}"].Value = subject.Seminars;
        //                    Total += subject.Seminars == null ? 0 : subject.Seminars.Value;
        //                    totalHours.Seminars += subject.Seminars ?? 0;
        //                    worksheet.Cells[$"K{row}"].Value = subject.Laboratory;
        //                    Total += subject.Laboratory == null ? 0 : subject.Laboratory.Value;
        //                    totalHours.Laboratory += subject.Laboratory ?? 0;
        //                    if (subject.SelfStudy != null)
        //                    {
        //                        double kons = Math.Round(((double)subject.SelfStudy * countFlow * 2.5 * 0.01), 1, MidpointRounding.AwayFromZero);
        //                        worksheet.Cells[$"L{row}"].Value = kons;
        //                        Total += kons;
        //                        totalHours.ConsultationsKSR += kons;
        //                    }

        //                    if (subject.ReportingForm == "зачет")
        //                    {
        //                        double zach = Math.Round(((double)subject.Commercial / 3), 1, MidpointRounding.AwayFromZero);
        //                        worksheet.Cells[$"O{row}"].Value = zach;
        //                        Total += zach;
        //                        totalHours.Test += zach;
        //                    }
        //                    else if (subject.ReportingForm == "экзамен")
        //                    {
        //                        double ekz = Math.Round(((double)subject.Commercial / 2), 1, MidpointRounding.AwayFromZero);
        //                        worksheet.Cells[$"M{row}"].Value = 2 * countFlow;
        //                        worksheet.Cells[$"N{row}"].Value = ekz;
        //                        Total += ekz + 2 * countFlow;
        //                        totalHours.Exams += ekz;
        //                        totalHours.Consultations += 2 * countFlow;
        //                    }

        //                    // Работы в ГАК
        //                    // ???????

        //                    // Проверка контрольных работ
        //                    double kr = Math.Round(((double)subject.Commercial / 6), 1, MidpointRounding.AwayFromZero);
        //                    worksheet.Cells[$"T{row}"].Value = kr;
        //                    Total += kr;
        //                    totalHours.ControlWorks += kr;

        //                    if (sub.Key.Trim() == "Руководство аспирантами")
        //                    {
        //                        worksheet.Cells[$"W{row}"].Value = subject.TotalHours;
        //                        Total += subject.TotalHours == null ? 0 : subject.TotalHours.Value;
        //                        totalHours.LeadershipGraduateStudents += subject.TotalHours ?? 0;
        //                    }

        //                    // Итого часов
        //                    worksheet.Cells[$"Z{row}"].Value = Total;
        //                    totalHours.Total += Total;
        //                }

        //                row++;
        //            }
        //        }
        //        //row++;
        //    }
        //    worksheet.Cells[$"A{row}"].Value = "Итого по " + faculty;
        //    totalHours.PrintToExcel(worksheet, row);
        //}
    }
}
