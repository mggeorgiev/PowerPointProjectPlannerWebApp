using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace PowerPointProjectPlannerWebApp.Models
{
    public class ProjectModel
    {
        [MaxLength(50), MinLength(3)]
        public string Title { get; set; }
        
        [MaxLength(50), MinLength(3), DataType("string")]
        public string Description { get; set; }

        public int Interval { get; set; }
        public int Duration { get; set; }
        public DateTime StartDate { get; set; }
    }
}
