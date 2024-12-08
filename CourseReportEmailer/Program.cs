using CourseReportEmailer.Models;
using CourseReportEmailer.Repository;
using CourseReportEmailer.Workers;
using Newtonsoft.Json;
using System.Data;

namespace CourseReportEmailer
{
    internal class Program
    {
        static void Main(string[] args)
        {

            try
            {
                EnrollmentDetailReportCommand command = new EnrollmentDetailReportCommand(@"Data Source=localhost;Initial Catalog=CourseReport;Integrated Security=True;Encrypt=False");
                IList<EnrollmentDetailReportModel> models = command.GetList(); //get the list from stored procedure

                var reportFileName = "EnrollmentDetailReport.xlsx"; //from requirements 

                EnrollmentDetailReportSpreadSheetCreator enrollmentSheetCreator = new EnrollmentDetailReportSpreadSheetCreator();
                enrollmentSheetCreator.Create(reportFileName, models); //here we are creating the sheets from given models

                EnrollmentDetailReportEmailSender emailer = new EnrollmentDetailReportEmailSender();
                emailer.Send(reportFileName);

            }
            catch (Exception ex)
            {
                Console.WriteLine("Something went wrong: {0}", ex.Message);
            }
            Console.ReadKey();
        }
    }
}
