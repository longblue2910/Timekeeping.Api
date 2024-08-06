using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Style;

namespace TT.API.Controllers;

[Route("api/[controller]")]
[ApiController]
public class TimekeepingController : ControllerBase
{

    [HttpPost]
    public async Task<IActionResult> UpdateCCCSAMAsync([FromForm] TimekeepingRequest request)
    {
        if (request.File.Length == 0) return NotFound();

        using (var stream = new MemoryStream())
        {
            // Copy file
            await request.File.CopyToAsync(stream);

            var fileName = request.File.FileName;

            // Creating an instance of ExcelPackage
            using (var excel = new ExcelPackage(stream))
            {
                // Get sheet excel
                ExcelWorksheet worksheet = excel.Workbook.Worksheets[0];
                var rowCount = worksheet.Dimension.Rows;

                if (rowCount == 0) return NotFound();

                List<TimekeepingResponse> responses = [];

                // Row index      
                for (int rowIndex = 3; rowIndex <= rowCount; rowIndex++)
                {
                    var employeeId = worksheet.Cells[$"A{rowIndex}"].Value?.ToString().Trim();
                    if (string.IsNullOrEmpty(employeeId)) continue;

                    var fn = worksheet.Cells[$"B{rowIndex}"].Value?.ToString().Trim();
                    if (string.IsNullOrEmpty(fn)) continue;

                    var department = worksheet.Cells[$"C{rowIndex}"].Value?.ToString().Trim();
                    if (string.IsNullOrEmpty(department)) continue;

                    var date = worksheet.Cells[$"D{rowIndex}"].Value?.ToString().Trim();
                    if (string.IsNullOrEmpty(date)) continue;

                    var startTimeStr = worksheet.Cells[$"F{rowIndex}"].Value?.ToString().Trim();
                    if (string.IsNullOrEmpty(startTimeStr)) continue;

                    var endTimeStr = worksheet.Cells[$"G{rowIndex}"].Value?.ToString().Trim();
                    if (string.IsNullOrEmpty(endTimeStr)) continue;

                    var dayOfWeek = worksheet.Cells[$"E{rowIndex}"].Value?.ToString().Trim();
                    if (string.IsNullOrEmpty(dayOfWeek)) continue;

                    if (TimeSpan.TryParse(startTimeStr, out TimeSpan startTime) && TimeSpan.TryParse(endTimeStr, out TimeSpan endTime))
                    {
                        // Assume both times are on the same day
                        DateTime startdate = DateTime.Today.Add(startTime);
                        DateTime enddate = DateTime.Today.Add(endTime);

                        // Calculate working hours and overtime
                        double workingHours = CalculateWorkingHours(startdate, enddate, dayOfWeek);
                        int overtimeHours = CalculateOvertimeHours(enddate);

                        DateTime officeEndDate = startdate.Date.AddHours(17).AddMinutes(0); // 17:00 AM
                        DateTime officeStartOTDate = startdate.Date.AddHours(21).AddMinutes(0); // 17:00 AM


                        if (workingHours < 0)
                        {
                            workingHours = 0;
                        }

                        if (startdate == enddate)
                        {
                            workingHours = 0;
                            overtimeHours = 0;
                        }

                        if (startdate > enddate)
                        {
                            workingHours = 0;
                            overtimeHours = 0;
                        }

                        if (startdate >= officeEndDate && startdate <= officeStartOTDate)
                        {
                            workingHours = 0;
                            overtimeHours = 0;
                        }



                        // Output results
                        Console.WriteLine($"Row {rowIndex}: Working time = {workingHours} hours, Overtime = {overtimeHours} hours");

                        double rounded = Math.Round(workingHours * 2, MidpointRounding.AwayFromZero) / 2;

                        var timeKeeping = new TimekeepingResponse
                        {
                            EmployeeId = employeeId,
                            FirstName = fn,
                            Department = department,
                            Date = date,
                            FirstPunch = startTimeStr,
                            LastPunch = endTimeStr,
                            Weekday = dayOfWeek,    
                            TotalTimeOT = overtimeHours,
                            TotalTimeWorking = rounded
                        };

                        responses.Add(timeKeeping); 
                    }
                }

                

                if (responses.Count != 0)
                {
                    var dirPath = Directory.GetCurrentDirectory();

                    var templateName = "template.xlsx";

                    var templatePath = $"{dirPath}/{templateName}";

                    // Creating an instance of ExcelPackage
                    var excelExport = new ExcelPackage(new FileInfo(templatePath), true);

                    //ExcelPackage.LicenseContext = LicenseContext.Commercial;

                    //Name of the sheet
                    var workSheet = excelExport.Workbook.Worksheets[0];

                    //Số dòng bắt đầu đổ data
                    int recordIndex = 2;
                    int stt = 1;
                    responses.ForEach(e =>
                    {

                        workSheet.Cells[$"A{recordIndex}:K{recordIndex}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        workSheet.Cells[$"A{recordIndex}:K{recordIndex}"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        workSheet.Cells[$"A{recordIndex}:K{recordIndex}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        workSheet.Cells[$"A{recordIndex}:K{recordIndex}"].Style.Border.Top.Style = ExcelBorderStyle.Thin;

                        workSheet.Row(recordIndex).Height = 20;
                        workSheet.Row(recordIndex).Style.WrapText = true;
                        workSheet.Row(recordIndex).Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        workSheet.Cells[$"A{recordIndex}"].Value = stt;
                        workSheet.Cells[$"A{recordIndex}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        workSheet.Cells[$"B{recordIndex}"].Value = e.FirstName;
                        workSheet.Cells[$"B{recordIndex}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        workSheet.Cells[$"C{recordIndex}"].Value = e.Department;
                        workSheet.Cells[$"C{recordIndex}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        workSheet.Cells[$"D{recordIndex}"].Value = e.Date;
                        workSheet.Cells[$"D{recordIndex}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        workSheet.Cells[$"E{recordIndex}"].Value = e.Weekday;
                        workSheet.Cells[$"E{recordIndex}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        workSheet.Cells[$"F{recordIndex}"].Value = e.FirstPunch;
                        workSheet.Cells[$"F{recordIndex}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        workSheet.Cells[$"G{recordIndex}"].Value = e.LastPunch;
                        workSheet.Cells[$"G{recordIndex}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        workSheet.Cells[$"H{recordIndex}"].Value = e.Weekday != "Sunday" ? e.TotalTimeWorking : null;
                        workSheet.Cells[$"H{recordIndex}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        workSheet.Cells[$"I{recordIndex}"].Value = e.Weekday != "Sunday" ? e.TotalTimeOT : null;
                        workSheet.Cells[$"I{recordIndex}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        workSheet.Cells[$"J{recordIndex}"].Value = e.Weekday == "Sunday" ? e.TotalTimeWorking : null;
                        workSheet.Cells[$"J{recordIndex}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        workSheet.Cells[$"K{recordIndex}"].Value = e.Weekday == "Sunday" ? e.TotalTimeOT : null;
                        workSheet.Cells[$"K{recordIndex}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        stt++;
                        recordIndex++;
                    });


                    var file = excelExport.GetAsByteArray();
                    return File(file, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"{fileName}");
                }
            }
        }

        return Ok();
    }

    private double CalculateWorkingHours(DateTime startdate, DateTime enddate, string dayOfWeek)
    {
        DateTime officeEndDate = startdate.Date.AddHours(17).AddMinutes(0); // 17:00 AM


        //DateTime earlyDate = startdate.Date.AddHours(7).AddMinutes(0); // 7:00 AM

        if (dayOfWeek != "Sunday")
        {
            DateTime officeStartDate = startdate.Date.AddHours(7).AddMinutes(30); // 7:30 AM

            DateTime lateDate0_5 = startdate.Date.AddHours(7).AddMinutes(45); // 7:45 AM

            DateTime lateDate1 = startdate.Date.AddHours(8).AddMinutes(0); // 8:00 AM

            DateTime lateDate2 = startdate.Date.AddHours(8).AddMinutes(30); // 8:30 AM

            DateTime lateDate3 = startdate.Date.AddHours(9).AddMinutes(00); // 8:30 AM


            if (startdate > officeStartDate && startdate < lateDate0_5)
            {
                startdate = officeStartDate;

            }

            if (startdate > lateDate0_5 && startdate < lateDate1)
            {
                startdate = lateDate1;
            }

            if (startdate > lateDate1 && startdate < lateDate2)
            {
                startdate = lateDate3;
            }

        }
        else
        {
            DateTime officeStartDate = startdate.Date.AddHours(8).AddMinutes(30); // 8:30 AM

            DateTime lateDate1 = startdate.Date.AddHours(8).AddMinutes(45); // 8:45 AM

            DateTime lateDate2 = startdate.Date.AddHours(9).AddMinutes(0); // 8:45 AM

            if (startdate >= officeStartDate & startdate < lateDate1)
            {
                startdate = officeStartDate;
            }

            if (startdate >= lateDate1 & startdate < lateDate2)
            {
                startdate = lateDate2;
            }

        }




        // Define lunch break times
        DateTime lunchStart = startdate.Date.AddHours(11).AddMinutes(30); // 11:30 AM
        DateTime lunchEnd = startdate.Date.AddHours(13); // 1:00 PM

        // Calculate total working time excluding lunch break
        TimeSpan workingTime = enddate - startdate;
        
        if (startdate < lunchEnd && enddate > lunchStart)
        {
            TimeSpan lunchBreak = lunchEnd - lunchStart;
            workingTime -= lunchBreak;
        }

        return workingTime.TotalHours;
    }

    private int CalculateOvertimeHours(DateTime enddate)
    {
        DateTime overtimeStart1 = enddate.Date.AddHours(20); // 8:00 PM
        DateTime overtimeStart2 = enddate.Date.AddHours(21); // 9:00 PM
        DateTime overtimeStart3 = enddate.Date.AddHours(22); // 10:00 PM
        DateTime overtimeStart4 = enddate.Date.AddHours(23); // 11:00 PM

        if (enddate > overtimeStart4)
        {
            return 4;
        }
        else if (enddate > overtimeStart3)
        {
            return 3;

        }

        else if (enddate > overtimeStart2)
        {
            return 2;
        }
        else if (enddate > overtimeStart1)
        {
            return 1;
        }

        return 0;
    }
}
public record TimekeepingRequest(IFormFile File);

public class TimekeepingResponse
{
    public string EmployeeId { get; set; }
    public string FirstName { get; set; }
    public string Department { get; set; }
    public string Date { get; set; }
    public string Weekday { get; set; }
    public string FirstPunch { get; set; }
    public string LastPunch { get; set; }
    public double? TotalTimeWorking { get; set; }
    public int TotalTimeOT { get; set; }

}
