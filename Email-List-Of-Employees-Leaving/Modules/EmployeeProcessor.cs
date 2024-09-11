using ClosedXML.Excel;
using EmployeeSendMailModels;
using System.Text.RegularExpressions;

namespace EmployeeSendMailProcessor
{
    public class EmployeeProcessor
    {
        DateTime tomorrow = DateTime.Today.AddDays(1);

        public void ProcessEmployees(Employee employeeResponse)
        {
            // Retrieve the list of employees from the API response
            List<Employee.Result> employees = employeeResponse.result;

            // Filter employees who resigned tomorrow
            List<Employee.Result> employeesLeavingTomorrow = employees?.Where(e => e.ngayNghiViec.HasValue && e.ngayNghiViec.Value.Date == tomorrow).ToList() ?? new List<Employee.Result>();

            // Exit the method if no employees resigned tomorrow
            if (!employeesLeavingTomorrow.Any())
            {
                Console.WriteLine("There are no employees leaving tomorrow!");
                return;
            }

            // Console number of employees leaving tomorrow
            Console.WriteLine($"{employeesLeavingTomorrow.Count} Employee{(employeesLeavingTomorrow.Count > 1 ? "s" : "")} leaving tomorrow!");

            // Create a new Excel workbook and worksheet
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Employees");

            // Define column headers
            string[] headers = new[]
            {
                    "Ordinal Number", "First Name", "Last Name", "Employee Code", "Email",
                    "Cost Center", "Cost Center Name", "Position Code", "Position Name",
                    "Probation Start Date", "Probation End Date", "Date Of Official Employment",
                    "Date Leave Work", "Status Code", "Status Name"
            };

            // Add headers to the worksheet and style them
            for (int i = 0; i < headers.Length; i++)
            {
                worksheet.Cell(2, i + 1).Value = headers[i];
            }

            // Populate worksheet with employee data
            int row = 3;
            foreach (var emp in employeesLeavingTomorrow)
            {
                worksheet.Cell(row, 1).Value = row - 2;
                worksheet.Cell(row, 2).Value = emp.firstName;
                worksheet.Cell(row, 3).Value = emp.lastName;
                worksheet.Cell(row, 4).Value = emp.employeeCode;
                worksheet.Cell(row, 5).Value = emp.email;
                worksheet.Cell(row, 6).Value = emp.costCenter;
                worksheet.Cell(row, 7).Value = emp.costCenterName;
                worksheet.Cell(row, 8).Value = emp.positionCode;
                worksheet.Cell(row, 9).Value = emp.positionName;
                worksheet.Cell(row, 10).Value = emp.ngayVaoLam?.ToString("yyyy-MM-dd") ?? "";
                worksheet.Cell(row, 11).Value = emp.ngayKetThucThuViec?.ToString("yyyy-MM-dd") ?? "";
                worksheet.Cell(row, 12).Value = emp.ngayLamChinhThuc?.ToString("yyyy-MM-dd") ?? "";
                worksheet.Cell(row, 13).Value = emp.ngayNghiViec?.ToString("yyyy-MM-dd") ?? "";
                worksheet.Cell(row, 14).Value = emp.maTinhTrang;
                worksheet.Cell(row, 15).Value = emp.tenTinhTrang;
                Console.WriteLine($"Full Name: {emp.firstName} {emp.lastName} | Employee Code: {emp.employeeCode} | Email: {emp.email} | Cost Center: {emp.costCenter}");
                row++;
            }

            // Style and format the worksheet
            var headerRange = worksheet.Range("A2:O2");
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            headerRange.Style.Font.FontSize = 13;

            worksheet.Range("A1:O1").Merge().Value = $"LIST OF EMPLOYEES LEAVING {tomorrow:yyyy-MM-dd}";
            worksheet.Cell(1, 1).Style.Font.FontSize = 15;
            worksheet.Cell(1, 1).Style.Fill.BackgroundColor = XLColor.FromArgb(197, 217, 241);
            worksheet.Row(1).Style.Font.Bold = true;

            worksheet.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Columns().AdjustToContents();
            worksheet.RangeUsed().Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            worksheet.RangeUsed().Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            // Save the workbook to a MemoryStream and send the email
            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            stream.Position = 0;
            SendMailToPersonInCharge(stream);
        }

        public void SendMailToPersonInCharge(MemoryStream fileStream)
        {
            SendMail mail = new SendMail();

            List<string> mailReceiver = GetMailReceivers();
            string fileName = $"List Of Employees Leaving {tomorrow:yyyy-MM-dd}.xlsx";

            // Check if mailReceiver is empty
            if (mailReceiver.Count == 0)
            {
                Console.WriteLine("No valid email address found. Email was not sent!");
                return;
            }

            // Create a new MemoryStream from the input fileStream
            MemoryStream fileStreamAttach = new MemoryStream(fileStream.ToArray());
            fileStreamAttach.Position = 0; // Reset the position of the stream

            // Define the email subject
            string mailSubject = $"List of employees leaving {tomorrow:yyyy-MM-dd}";

            // Create the email body with HTML content
            string mailHTML = $@"
                                 <b>Dear Everyone</b><br>
                                 <p>This is the list of employees leaving tomorrow.</p>
                                ";

            // Send the email with the attached MemoryStream file
            mail.SendMailEmployeesLeavingTomorrow(mailReceiver, mailSubject, mailHTML, fileStreamAttach, fileName);

            Console.WriteLine("Email sent successfully to valid emails!");
        }

        public static List<string> GetMailReceivers()
        {
            string folderPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + '\\';
            var mailReceivers = new List<string>();

            using (var workbook = new XLWorkbook(folderPath + "List Send Mail For PIC.xlsx"))
            {
                var worksheet = workbook.Worksheet(1);
                foreach (var row in worksheet.RowsUsed().Skip(2))
                {
                    var email = row.Cell(3).GetValue<string>();
                    if (!string.IsNullOrWhiteSpace(email) && IsValidEmail(email))
                    {
                        mailReceivers.Add(email);
                    }
                }
            }
            return mailReceivers;
        }

        static bool IsValidEmail(string email)
        {
            string pattern = @"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$";
            return Regex.IsMatch(email, pattern);
        }
    }
}
