using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

public static class FillOverTimeService
{
    public static IWorkbook Execute(Dictionary<string, string> calculateOverTime, IWorkbook workbook)
    {

        // Create a new worksheet in the output workbook
        ISheet outputSheet = workbook.GetSheet("Output");

        var size = outputSheet.LastRowNum - 1;
        // Iterate over the rows in the input worksheet
        for (int i = 0; i <= size; i++)
        {
            // Create a new row in the output worksheet
            IRow outputRow = outputSheet.GetRow(i + 1);

            var perCell = outputRow.GetCell(0).ToString();

            // Extract overtime map to specific personal number
            var overTime = calculateOverTime.ContainsKey(perCell ?? "") ? calculateOverTime[perCell ?? ""] : string.Empty;

            // Create cells in the output row and write the values
            outputRow.CreateCell(6).SetCellValue("overTime");

        }

        return workbook;

    }
}

