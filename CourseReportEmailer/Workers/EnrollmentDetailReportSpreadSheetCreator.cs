using CourseReportEmailer.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CourseReportEmailer.Workers
{
    internal class EnrollmentDetailReportSpreadSheetCreator
    {
        public void Create(string fileName, IList<EnrollmentDetailReportModel> enrollmentModels)
        {
            //specifying the name and type of the document to create
            using(SpreadsheetDocument document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                //do the JSON trick to convert EnrollmentDetailReportModel into DataTable
                var json = JsonConvert.SerializeObject(enrollmentModels); //a list of enrollment models

                //next we need to put that into a DataTable 
                DataTable enrollmentsTable = (DataTable) JsonConvert.DeserializeObject(json, typeof(DataTable));

                //there are several components - Workbook, WorkSheet, Sheets
                //A Sheet can be any type (chart sheet, work sheet etc) in excel it's like a parent object
                //Workbook is the entire document, it can have different sheets and those sheets can be chart sheets or work sheets

                //first we need to create a workbook
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                //WorkbookPart is an object that contains global settings about all of the components in a workbook
                //and then the Workbook is the actual workbook
                //so we created a WorkbookPart and within that we created the Workbook

                //next we need to create a worksheet
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>(); //from workbookPart we want to add a new part to the workbook
                                                                                        //and that's a worksheetPart

                //WorksheetPart is a global setting about the worksheet

                //there is an object called sheet data and sheet data is the actual data on the worksheet
                //this is where we want to put our tables and everything else
                //this is where you're going to visually see what's on the worksheet
                SheetData sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData); //worksheer takes another object which is sheetData
                                                                    //and we we will use sheetData to actually add the table
                                                                    //like a lego: sheetData <- table is here (what you visually see),
                                                                    //sheetData is in worksheet object which is in worksheetPart

                //next we need to associate our worksheetPart to our workbookPart, because so far it's just a separate worksheet and a separate workbook
                Sheets sheetList = document.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets()); // here we are saying that our workbook
                                                                                                     // is gonna have a list of generic sheets
                                                                                                     // (can be work sheets that we have above
                                                                                                     // or it can be chart sheets)
                                                                                                     //now workbook knows it has a list of sheets,
                                                                                                     //but so far they are empty

                //let's add a generic sheet to our list of sheets
                Sheet singleSheet = new Sheet()
                {
                    //we still don't know the connection between worksheet and workbook => here we create the connection between them
                    Id = document.WorkbookPart.GetIdOfPart(worksheetPart), //this method is going to go and grab our worksheet's ID
                                                                          //which is some ID automatically assigned to it
                                                                          //and it's going to say this sheet is basically going to be 
                                                                          //this worksheet (singleSheet) right here (worksheetPart)
                                                                          //that it's going to be associated with workbook (WorkbookPart)
                                                                          //this is how we associate the work sheet with our workbook

                    SheetId = 1,
                    //part of the requirement - worksheet has to be named report sheet
                    Name = "Report Sheet"
                };
                sheetList.Append(singleSheet); //add that sheet to the list which belongs to the workbook
                                               //and that's how the workbook knows which worksheet to use
                                               //you can similarly add more sheets here to the list

                //next we put the data to the table on the worksheet

                //first let's create title columns
                Row excelTitleRow = new Row(); //the Row object comes from DocumentFormat.OpenXml.Spreadsheet

                //go through columns in enrollmentsTable to create a header
                foreach (DataColumn tableColumn in enrollmentsTable.Columns)
                {
                    Cell cell = new Cell(); //create a cell object that comes from DocumentFormat.OpenXml.Spreadsheet
                    cell.DataType = CellValues.String; //we are sayng what type of the cell will be
                    cell.CellValue = new CellValue(tableColumn.ColumnName); //what cell value it will be => taking the value from the column name

                    excelTitleRow.Append(cell); //add that cell to our excel title row (= our header row) like First name, Last name, Id, etc
                }

                //add a new row on that sheet - sheetData has all data
                sheetData.AppendChild(excelTitleRow); //add the first row that we will see (title row)

                //next we need to add an actual info
                foreach (DataRow tableRow in enrollmentsTable.Rows)
                {
                    //grad the row
                    Row excelNewRow = new Row();
                    foreach (DataColumn tableColumn in enrollmentsTable.Columns)
                    {
                        //here we try to find an intersection between row and column so that we can find the cell value
                        Cell cell = new Cell();
                        cell.DataType = CellValues.String;
                        cell.CellValue = new CellValue(tableRow[tableColumn.ColumnName].ToString());
                        excelNewRow.AppendChild(cell);
                    }

                    //each row needs to be put in sheet data 
                    sheetData.AppendChild(excelNewRow);
                }

                workbookPart.Workbook.Save(); //it will use the file name that we gave it originally and it's gonna save our workbook
            }
        }
    }
}
