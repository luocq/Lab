using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.Entity;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
namespace CodeFirstDemo.Models
{
    public class ProjectSchedule
    {
        public Guid ID { get; set; }

        public DateTime StartTime { get; set; }

        public DateTime EndTime { get; set; }

        public Guid ProjectID { get; set; }

        [ForeignKey("ProjectID")]
        public virtual Project Project { get; set; }
    }
}