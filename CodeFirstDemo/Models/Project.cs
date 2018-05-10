using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CodeFirstDemo.Models
{
    public class Project
    {
        public Project() {
            this.ID = Guid.NewGuid();
            this.ProjectSchedules = new List<ProjectSchedule>();
        }
        public Guid ID { get; set; }

        public string Name { get; set; }

        public virtual ICollection<ProjectSchedule> ProjectSchedules { get; set; }
    }
}