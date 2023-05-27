namespace WebSSU.Models
{
    public class TimeNorms
    {
        // производственная практика
        public int ProductionPractice { get; set; }
        // учебная практика (НИР)
        public int EducationalPractice { get; set; }
        //  Норма часов по курсовым работам (2 курс)
        public int Coursework2 { get; set; }
        // Норма часов по курсовым работам (3 курс)
        public int Coursework3 { get; set; }
        // Норма часов по курсовым работам для ПО (2 курс)
        public int Coursework2PO { get; set; }
        // Норма часов по курсовым работам для ПО (3 курс)
        public int Coursework3PO { get; set; }
        // ВКР (бакалавриат)
        public int FinalQualifyingWorkBachelor { get; set; }
        // ВКР (магистратура)
        public int FinalQualifyingWorkMagistracy { get; set; }

    }
}
