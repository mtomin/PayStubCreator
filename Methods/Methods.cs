using PayStubCreator;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;
using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Reflection;

namespace Methods
{
    public class EventData : EventArgs
    {
        //Container for error description
        public string ErrorDescription { get; set; }
    }
    public class Readers
    {
        static public event EventHandler<EventData> ErrorOccured;
        static public EventData eventData=new EventData();
        
        private static void RaiseError(string errorDescription)
        {
            if (ErrorOccured != null)
            {
                eventData.ErrorDescription = errorDescription;
                ErrorOccured(null, eventData);
            }
        }

        public static void ReadCompanyData(Company company, OfficeOpenXml.ExcelWorksheet worksheet)
        {
            //populate Company company
            try
            {
                company.Name = worksheet.Cells["B2"].Text;
                company.Address = worksheet.Cells["C3"].Text;
                company.City = worksheet.Cells["C5"].Text.Split(',')[1];
                company.PostCode = worksheet.Cells["C5"].Text.Split(',')[0];
            }
            catch
            {
                RaiseError("An error was encountered while parsing the company data in the file");
            }
            foreach (PropertyInfo property in typeof(Company).GetProperties())
            {
                if ((string)property.GetValue(company)=="")
                {
                    RaiseError("Some or all company data appears to be missing in file");
                    break;
                }
            }
        }

        public static void GetEmployeeData(Employee employee, OfficeOpenXml.ExcelWorksheet worksheet)
        {
            //Try populating Employee employee with data from the excel file. Raise an error if not possible.
            try
            {
                employee.FirstName = worksheet.Cells["C6"].Text.Split(' ')[0];
                employee.LastName = worksheet.Cells["C6"].Text.Split(' ')[1];
                employee.EmployeeID = int.Parse(worksheet.Cells["C7"].Text);
                employee.HoursWorked = int.Parse(worksheet.Cells["D24"].Text);
                employee.HoursOvertime = int.Parse(worksheet.Cells["E24"].Text);
                employee.HoursVacation = int.Parse(worksheet.Cells["G24"].Text);
                employee.SickLeave = int.Parse(worksheet.Cells["F24"].Text);

                foreach (var cell in worksheet.Cells["H10:H23"])
                {
                    if (cell.Text != "0")
                    {
                        employee.DaysWorked++;
                    }
                }
            }
            catch
            {
                RaiseError("An error was encountered while parsing the employee data in the file");
            }
            if (employee.FirstName=="" || employee.LastName =="")
            {
                RaiseError("Employee data appears to be missing in file");
            }
        }

        public static void LoadDBInfo(Employee employee)
        //Loads user info from the EmployeeData database
        {
            string connectionString = ConfigurationManager.ConnectionStrings["EmployeeData"].ConnectionString;
            SqlConnection connection;
            DataTable EmployeeData = new DataTable();
            string query = "SELECT " +
                "Address, City, PostCode, Title, Coefficient, DistanceFromWork, NumberOfDependents " +
                "FROM Employees " +
                "WHERE FirstName=@employeeFirstName AND LastName=@employeeLastName AND EmployeeID=@employeeID";

            using (connection = new SqlConnection(connectionString))
            using (SqlCommand command = new SqlCommand(query, connection))
            using (SqlDataAdapter adapter = new SqlDataAdapter(command))
            {
                command.Parameters.AddWithValue("@employeeFirstName", employee.FirstName);
                command.Parameters.AddWithValue("@employeeLastName", employee.LastName);
                command.Parameters.AddWithValue("@employeeID", employee.EmployeeID);

                adapter.Fill(EmployeeData);
            }
            
            //Check if employee exists
            if (EmployeeData.Rows.Count == 0)
            {
                RaiseError(String.Format("The employee {0} {1} with ID {2} was not found in database. Please double-check the information provided in file", employee.FirstName, employee.LastName, employee.EmployeeID));
            }

            //fill all properties from the employee database if possible. If not, raise an error.
            try
            {
                employee.Address = EmployeeData.Rows[0]["Address"].ToString();
                employee.City = EmployeeData.Rows[0]["City"].ToString();
                employee.PostCode = (int)EmployeeData.Rows[0]["PostCode"];
                employee.Title = EmployeeData.Rows[0]["Title"].ToString();
                employee.Coefficient = (decimal)EmployeeData.Rows[0]["Coefficient"];
                employee.DistanceFromWork = (decimal)EmployeeData.Rows[0]["DistanceFromWork"];
            }
            catch
            {
                String.Format("Some or all data for the employee {0} {1} with ID {2} could not be read from database.", employee.FirstName, employee.LastName, employee.EmployeeID);
            }
            
            if (EmployeeData.Rows[0]["NumberOfDependents"] == DBNull.Value)
            {
                employee.NumberOfDependents = null;
            }
            else
            {
                employee.NumberOfDependents = (int)EmployeeData.Rows[0]["NumberOfDependents"];
            }

            if (employee.Address=="" || employee.City=="" || employee.City == "" || employee.City == "Title")
            {
                RaiseError(String.Format("Some or all data for the employee {0} {1} with ID {2} is missing from database.", employee.FirstName, employee.LastName, employee.EmployeeID));
            }
        }
    }
    public class Writers
    {
        public static void ExportToPDF(Company company, Employee employee, string outputFilePath)
        {
            //Page formatting constants
            const int HEADER_WIDTH = 250;
            const int HEADER_HEIGHT = 100;
            const int LEFT_MARGIN = 40;
            const int BODY_WIDTH = 380;
            const int BODY_HEIGHT = 100;
            double currentHeight = 50;

            PdfDocument pdf = new PdfDocument();
            PdfPage pdfPage = pdf.AddPage();
            XGraphics graph = XGraphics.FromPdfPage(pdfPage);
            XTextFormatter tf = new XTextFormatter(graph);

            //Header section

            XFont headerFont = new XFont("Times New Roman", 16, XFontStyle.Bold);
            tf.Alignment = XParagraphAlignment.Left;
            tf.DrawString(company.Header(), headerFont, XBrushes.Black, new XRect(LEFT_MARGIN, currentHeight, HEADER_WIDTH, HEADER_HEIGHT), XStringFormats.TopLeft);
            tf.Alignment = XParagraphAlignment.Right;
            currentHeight += HEADER_HEIGHT;
            tf.DrawString(employee.AddressHeader(), headerFont, XBrushes.Black, new XRect(pdfPage.Width - HEADER_WIDTH - LEFT_MARGIN, currentHeight, HEADER_WIDTH, HEADER_HEIGHT), XStringFormats.TopLeft);
            string companyLogoPath = @"C:\Users\Menta\Desktop\mfclogo.png";
            XImage companyLogo = XImage.FromFile(companyLogoPath);
            graph.DrawImage(companyLogo, LEFT_MARGIN, currentHeight - 30, HEADER_WIDTH, HEADER_HEIGHT);

            //Body section

            //Monospaced font - kerning issues
            XFont bodyFont = new XFont("Consolas", 12);

            tf.Alignment = XParagraphAlignment.Left;
            string payBreakdown = "";
            payBreakdown += "Regular hours worked:" + employee.HoursWorked.ToString().PadLeft(50 - "Regular hours worked:".Length) + "\n";
            payBreakdown += "Overtime hours worked:" + employee.HoursOvertime.ToString().PadLeft(50 - "Overtime hours worked:".Length) + "\n";
            payBreakdown += "Vacation hours used:" + employee.HoursVacation.ToString().PadLeft(50 - "Vacation hours used:".Length) + "\n";
            payBreakdown += "Sick leave hours:" + employee.SickLeave.ToString().PadLeft(50 - "Sick leave hours:".Length) + "\n";
            payBreakdown += "Days worked:" + employee.DaysWorked.ToString().PadLeft(50 - "Days worked:".Length) + "\n";
            payBreakdown += "Employee coefficient:" + employee.Coefficient.ToString().PadLeft(50 - "Employee coefficient:".Length);

            tf.DrawString(payBreakdown, bodyFont, XBrushes.Black, new XRect(LEFT_MARGIN, pdfPage.Height / 3, BODY_WIDTH, BODY_HEIGHT), XStringFormats.TopLeft);
            currentHeight = pdfPage.Height / 3;

            //increase the height for the amount of rows the paybreakdown string has * height of font 12
            currentHeight += graph.MeasureString(payBreakdown, new XFont("Consolas", 12)).Height * payBreakdown.Split('\n').Length;

            //Draw horizontal line
            XPen separatorLine = new XPen(XColors.Black, 1);
            graph.DrawLine(separatorLine, LEFT_MARGIN, currentHeight, BODY_WIDTH, currentHeight);
            currentHeight += 10;

            string payItemization = "Total pay: " + String.Format("{0:F2}", employee.TotalPay()).PadLeft(50 - "Total pay: ".Length) + "\n\n";
            payItemization += "Deductions:\n\n";

            payItemization += "Income tax:" + employee.TaxesOwed().ToString().PadLeft(50 - "Income tax:".Length) + "\n";
            payItemization += $"Retirement deduction ({Employee.RETIREMENT_CONTRIBUTION_RATE * 100}%):" + String.Format("{0:F2}", employee.RetirementDeduction()).PadLeft(50 - $"Retirement deduction ({Employee.RETIREMENT_CONTRIBUTION_RATE * 100}%):".Length) + "\n";

            string travelExpenses = "Travel expenses ";
            travelExpenses += (employee.DistanceFromWork <= 10) ? "(public transit ticket):" : "(private vehicle):";
            travelExpenses += string.Format("{0:F2}", employee.TravelExpense()).PadLeft(50 - travelExpenses.Length);

            payItemization += travelExpenses;
            tf.DrawString(payItemization, bodyFont, XBrushes.Black, new XRect(LEFT_MARGIN, currentHeight, BODY_WIDTH, BODY_HEIGHT), XStringFormats.TopLeft);

            //Summary section
            string netPay = "Take home pay:";
            netPay += string.Format("{0:F2}", employee.NetPay()).PadLeft(50 - netPay.Length);
            currentHeight = 2 * pdfPage.Height / 3;
            tf.DrawString(netPay, bodyFont, XBrushes.Black, new XRect(LEFT_MARGIN, currentHeight, BODY_WIDTH, BODY_HEIGHT), XStringFormats.TopLeft);
            currentHeight += 50;
            graph.DrawLine(separatorLine, LEFT_MARGIN, currentHeight, LEFT_MARGIN + 200, currentHeight);
            graph.DrawLine(separatorLine, LEFT_MARGIN + 250, currentHeight, LEFT_MARGIN + 400, currentHeight);
            tf.DrawString(employee.FullName, bodyFont, XBrushes.Black, new XRect(LEFT_MARGIN, currentHeight, BODY_WIDTH, BODY_HEIGHT), XStringFormats.TopLeft);
            tf.DrawString(string.Format("for {0}", company.Name), bodyFont, XBrushes.Black, new XRect(LEFT_MARGIN + 250, currentHeight, BODY_WIDTH, BODY_HEIGHT), XStringFormats.TopLeft);

            outputFilePath = String.Format("{0}_{1}_{2}_{3}{4}", System.IO.Path.ChangeExtension(outputFilePath, null), employee.FirstName, employee.LastName, employee.EmployeeID, System.IO.Path.GetExtension(outputFilePath));
            pdf.Save(outputFilePath);
        }
    }
}
