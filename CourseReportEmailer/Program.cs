using CourseReportEmailer.Models;
using Newtonsoft.Json;

namespace CourseReportEmailer
{
    internal class Program
    {
        static void Main(string[] args)
        {
            EnrollmentDetailReportModel model = new EnrollmentDetailReportModel()
            {
                EnrollmentId = 1,
                FirstName = "Mark",
                LastName = "Hue",
                CourseCode = "CA",
                Description = "description",
            };

            //turn C# object into JSON object (= serialize)
            var json = JsonConvert.SerializeObject(model);

            //turn JSON object back to C# object
            EnrollmentDetailReportModel objectFromJson = (EnrollmentDetailReportModel) JsonConvert.DeserializeObject(json, typeof(EnrollmentDetailReportModel));

            Console.ReadKey();
        }
    }
}
