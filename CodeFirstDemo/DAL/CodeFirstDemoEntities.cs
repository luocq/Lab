using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace CodeFirstDemo.DAL
{
    public class CodeFirstDemoEntities:DbContext
    {
        public DbSet<Models.Project> Projects { get; set; }
        public DbSet<Models.ProjectSchedule> ProjectSchedules { get; set; }
    }
}