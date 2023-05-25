using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace WebSSU.Models
{
    public class Subject
    {
        public string Name { get; set; }
        public string? Specialization { get; set; }
        public int? Semester { get; set; }
        public int? Budget { get; set; }
        public int? Commercial { get; set; }
        public string? Groups { get; set; }
        public string GroupForm { get; set; }
        public int? TotalHours { get; set; }
        public int? Lectures { get; set; }
        public int? Seminars { get; set; }
        public int? Laboratory { get; set; }
        public int? SelfStudy { get; set; }
        public int? LoadPerWeek { get; set; }
        public string? ReportingForm { get; set; }
        public string? Remark { get; set; }
    }
}
