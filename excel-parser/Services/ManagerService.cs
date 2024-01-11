using Serilog;
using Serilog.Core;
using Serilog.Events;
using Microsoft.Extensions.Logging;
using ILogger = Serilog.ILogger;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

public class ManagerService(ILogger<ManagerService> logger)
{
    private readonly ILogger<ManagerService> logger = logger;
    public readonly List<OuutputModel>? ouutputModels;

    public void Execute(string perNumAndwrGroupfilePath, string mainFilePath, string holidaysWorkFilePath, string specialFilePath)
    {
        this.logger.LogDebug("Manager service started");

        // Create empty workbook with spicific headers
        var workbook = InitialEmptyExcell.Execute();

        // Fill personal number and name to output excell
        var fillperNumAndwrGroupWorkGroup = FillPerNumberAndName.Execute(perNumAndwrGroupfilePath, workbook);

        // Fill zero value for spicifig columns
        var fillZzervoValues = FillZeroCell.Execute(fillperNumAndwrGroupWorkGroup);

        // Save the workbook to a file
        using (FileStream stream = new FileStream("Output.xlsx", FileMode.Create, FileAccess.Write))
        {
            fillZzervoValues.Write(stream);
        }

    }
}
