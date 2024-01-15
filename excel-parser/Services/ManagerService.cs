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

        // Create empty workbook with spicific headers
        var workbook = InitialEmptyExcell.Execute();

        // Fill personal number to output excell
        //var fillperNumAndwrGroupWorkGroup = FillPerNumber.Execute(perNumAndwrGroupfilePath, workbook);

        // Fill work group dictionary
        var workGroupDic = FiilWorkGroupDic.Execute(perNumAndwrGroupfilePath);

        // Calculate overtime
        var calculateOverTime = OverTimeCalculate.Excute(workGroupDic, mainFilePath);

        // Fill over time
       // var fillOverTime = FillOverTimeService.Execute(calculateOverTime, fillperNumAndwrGroupWorkGroup);

        // Fill zero value for spicifig columns
        // var fillZzervoValues = FillZeroCell.Execute(fillOverTime);

        // // Fill DailyMission, FractionOfWorkAbcenc, TotalFinancialFunction
        // var ProcessOfSpecialExcell = ProcessorOfSpecialExcell.Execute(specialFilePath, fillZzervoValues);

        // // Fill WorkingHolidays and WorkingAtNight fields
        // var ProcessOfNightHolidayExcell = NightHolidayWorkService.Execute(holidaysWorkFilePath, ProcessOfSpecialExcell);

        // // Save the workbook to a file
        // using (FileStream stream = new FileStream("Output.xlsx", FileMode.Create, FileAccess.Write))
        // {
        //     ProcessOfNightHolidayExcell.Write(stream);
        // }

    }
}
