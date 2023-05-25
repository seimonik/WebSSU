using OfficeOpenXml;

namespace WebSSU.Models
{
    public class TotalHours
    {
        public double Lectures { get; set; }
        public double Seminars { get; set; }
        public double Laboratory { get; set; }
        // Консультацтации по дисциплине, КСР
        public double ConsultationsKSR { get; set; }
        // Консультации перед экзаменом
        public double Consultations { get; set; }
        // Экзамены
        public double Exams { get; set; }
        // Зачкты
        public double Test { get; set; }
        // Руководство практикой
        public double Practice { get; set; }
        // Курсовые работы
        public double Coursework { get; set; }
        // ВКР
        public double VKR { get; set; }
        // работа в ГАК
        public double GAK { get; set; }
        // проверка контрольных работ
        public double ControlWorks { get; set; }
        // руководство аспирантами
        public double LeadershipGraduateStudents { get; set; }
        // руководство соискателями
        public double LeadershipApplicants { get; set; }
        // руководство магистерской программой
        public double LeadershipMaster { get; set; }
        // Факультативные занятия
        public double OptionalClasses { get; set; }
        public double Total { get; set; }
        public TotalHours()
        {
            Lectures = 0;
            Seminars = 0;
            Laboratory = 0;
            ConsultationsKSR = 0;
            Consultations = 0;
            Exams = 0;
            Test = 0;
            Practice = 0;
            Coursework = 0;
            VKR = 0;
            GAK = 0;
            ControlWorks = 0;
            LeadershipGraduateStudents = 0;
            LeadershipApplicants = 0;
            LeadershipMaster = 0;
            OptionalClasses = 0;
            Total = 0;
        }
        public void PrintToExcel(ExcelWorksheet worksheet, int row)
        {
            worksheet.Cells[$"I{row}"].Value = Lectures;
            worksheet.Cells[$"J{row}"].Value = Seminars;
            worksheet.Cells[$"K{row}"].Value = Laboratory;
            worksheet.Cells[$"L{row}"].Value = ConsultationsKSR;
            worksheet.Cells[$"M{row}"].Value = Consultations;
            worksheet.Cells[$"N{row}"].Value = Exams;
            worksheet.Cells[$"O{row}"].Value = Test;
            worksheet.Cells[$"P{row}"].Value = Practice;
            worksheet.Cells[$"Q{row}"].Value = Coursework;
            worksheet.Cells[$"R{row}"].Value = VKR;
            worksheet.Cells[$"S{row}"].Value = GAK;
            worksheet.Cells[$"T{row}"].Value = ControlWorks;
            worksheet.Cells[$"U{row}"].Value = LeadershipGraduateStudents;
            worksheet.Cells[$"V{row}"].Value = LeadershipApplicants;
            worksheet.Cells[$"W{row}"].Value = LeadershipMaster;
            worksheet.Cells[$"X{row}"].Value = OptionalClasses;
            worksheet.Cells[$"Z{row}"].Value = Total;
        }
    }
}
