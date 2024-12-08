using CourseReportEmailer.Models;
using Newtonsoft.Json;
using System.Data;

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

            //practice DataTable 
            DataTable table = new DataTable();

            //"EnrollmentId' is the name of the column etc
            DataColumn column1 = new DataColumn("EnrollmentId", typeof(int));
            DataColumn column2 = new DataColumn("FirstName", typeof(string));
            DataColumn column3 = new DataColumn("LastName", typeof(string));
            DataColumn column4 = new DataColumn("CourseCode", typeof(string));
            DataColumn column5 = new DataColumn("Description", typeof(string));

            //add these columns to the table 
            table.Columns.Add(column1);
            table.Columns.Add(column2);
            table.Columns.Add(column3);
            table.Columns.Add(column4);
            table.Columns.Add(column5);

            //put only one row to the table 
            table.Rows.Add(1, "Mark", "Hue", "CA", "description");

            //iterate this table
            foreach (DataRow row in table.Rows)
            {
                //now we go through the data columns
                foreach (DataColumn column in table.Columns)
                {
                    //on the intersection of row and column you will get the value
                    //Note - need to be attentive with the types otherwise results can be unexpected as it would try to convert to the original types
                    Console.WriteLine(row[column]);
                }

            }

            //why do we want to convert JSON into DataTable? => because it would be easier to use the data in worksheet this way

            Console.ReadKey();
        }
    }
}
