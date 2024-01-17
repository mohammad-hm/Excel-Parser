using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

public static class FillZeroCell
{
    public static IWorkbook Execute(IWorkbook workbook)
    {

        // Create a new worksheet in the output workbook
        ISheet outputSheet = workbook.GetSheet("Output");


        // Iterate over the rows in the input worksheet
        for (int i = 1; i <= outputSheet.LastRowNum; i++)
        {
            // Get the current row
            IRow outputRow = outputSheet.GetRow(i);

            // Get the spicifi cell in the output row
            ICell outputCell3 = outputRow.CreateCell(3);
            // Set the value of the spicific cell to zero
            outputCell3.SetCellValue("000:00");
            // Get the spicifi cell in the output row
            ICell outputCell8 = outputRow.CreateCell(8);
            // Set the value of the spicific cell to zero
            outputCell8.SetCellValue("000:00");
             // Get the spicifi cell in the output row
            ICell outputCell9 = outputRow.CreateCell(9);
            // Set the value of the spicific cell to zero
            outputCell9.SetCellValue("000:00");
             // Get the spicifi cell in the output row
            ICell outputCell12 = outputRow.CreateCell(12);
            // Set the value of the spicific cell to zero
            outputCell12.SetCellValue("000:00");
        }

        return workbook;

    }
}

