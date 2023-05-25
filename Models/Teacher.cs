using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace WebSSU.Models
{
    public class Teacher
    {
        public string? Name { get; set; }
        public List<Subject> subjects { get; set; }
        public Teacher()
        {
            subjects = new List<Subject>();
        }
    }
}
