using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CourseReportEmailer.Models
{
    internal class EnrollmentDetailReport
    {
        public int EnrollmentId { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string CourseCode { get; set; }
        public string Description { get; set; }
    }
}
