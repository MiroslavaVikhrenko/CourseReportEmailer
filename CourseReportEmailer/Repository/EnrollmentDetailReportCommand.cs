using CourseReportEmailer.Models;
using Dapper;
using Microsoft.Data.SqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CourseReportEmailer.Repository
{
    internal class EnrollmentDetailReportCommand
    {
        private string _connectionString;
        public EnrollmentDetailReportCommand(string connectionString)
        {
            _connectionString = connectionString;
        }

        public IList<EnrollmentDetailReportModel> GetList()
        {
            List<EnrollmentDetailReportModel> enrollmentDetailReport = new List<EnrollmentDetailReportModel>();

            //create store procedure using dapper
            var sql = "EnrollmentReport_GetList";

            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                enrollmentDetailReport = connection.Query<EnrollmentDetailReportModel>(sql).ToList();
            }

            return enrollmentDetailReport;
        }
    }
}
