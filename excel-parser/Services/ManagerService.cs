using Serilog;
using Serilog.Core;
using Serilog.Events;
using Microsoft.Extensions.Logging;
using ILogger = Serilog.ILogger;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using excel_parser.Services;

public class ManagerService(ILogger<ManagerService> logger)
{
    private readonly ILogger<ManagerService> logger = logger;

    public void Execute(string perNumAndwrGroupfilePath, string mainFilePath, string holidaysWorkFilePath, string specialFilePath)
    {
        this.logger.LogDebug("Manager service started");

        // Create empty workbook with specific headers
        var workbook = InitialEmptyExcell.Execute();
        // Save the workbook to a file
        using (FileStream stream = new FileStream("InitialEmptyExcell.xlsx", FileMode.Create, FileAccess.Write))
        {
            workbook.Write(stream);
        }

        // Fill personal number to output excel
        var fillperNumAndwrGroupWorkGroup = FillPerNumber.Execute(perNumAndwrGroupfilePath, workbook);
        using (FileStream stream = new FileStream("fillperNumAndwrGroupWorkGroup.xlsx", FileMode.Create, FileAccess.Write))
        {
            fillperNumAndwrGroupWorkGroup.Write(stream);
        }

        // Fill work group dictionary
        var workGroupDic = FiilWorkGroupDic.Execute(perNumAndwrGroupfilePath, fillperNumAndwrGroupWorkGroup);

        // Calculate overtime - Fill a dictionary that have personal number and overtime
        var calculateOverTime = OverTimeCalculate.Excute(workGroupDic, mainFilePath);
        
        // Fill over time
        var fillOverTime = FillOverTimeService.Execute(calculateOverTime, fillperNumAndwrGroupWorkGroup);
        using (FileStream stream = new FileStream("fillOverTime.xlsx", FileMode.Create, FileAccess.Write))
        {
            fillOverTime.Write(stream);
        }

        // Fill zero value for specific columns
        var fillZzervoValues = FillZeroCell.Execute(fillOverTime);
        using (FileStream stream = new FileStream("fillZzervoValues.xlsx", FileMode.Create, FileAccess.Write))
        {
            fillZzervoValues.Write(stream);
        }

        // Fill DailyMission, FractionOfWorkAbcenc, TotalFinancialFunction
        var ProcessOfSpecialExcell = ProcessorOfSpecialExcell.Execute(specialFilePath, fillZzervoValues);
        using (FileStream stream = new FileStream("ProcessOfSpecialExcell.xlsx", FileMode.Create, FileAccess.Write))
        {
            ProcessOfSpecialExcell.Write(stream);
        }


        // Fill WorkingHolidays and WorkingAtNight fields
        var ProcessOfNightHolidayExcell = NightHolidayWorkService.Execute(holidaysWorkFilePath, ProcessOfSpecialExcell);
       

        // Save the workbook to a file
        using (FileStream stream = new FileStream("ProcessOfNightHolidayExcell.xlsx", FileMode.Create, FileAccess.Write))
        {
            ProcessOfNightHolidayExcell.Write(stream);
        }

    }
}
