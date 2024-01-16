using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

public static class FillOverTimeService
{
    public static IWorkbook Execute(Dictionary<string, string> calculateOverTime, IWorkbook workbook)
    {
        
        foreach (var key in calculateOverTime.Keys.ToList())
        {
            if (int.TryParse(calculateOverTime[key], out var value))
            {
                // Check if the value is negative, replace with "0"
                if (value < 0)
                {
                    calculateOverTime[key] = "0";
                }
            }
            else
            {
                // Handle the case where the value is not a valid integer
                calculateOverTime[key] = "0";
            }
        }

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

