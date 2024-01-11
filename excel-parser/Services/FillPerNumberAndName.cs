using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

public static class FillPerNumberAndName
{
    public static IWorkbook Execute(string perNumAndwrGroupfilePath, IWorkbook workbook)
    {


        // Create a workbook object from the input file
        IWorkbook inputWorkbook = WorkbookFactory.Create(perNumAndwrGroupfilePath);
        // Get the first worksheet in the input workbook
        ISheet inputSheet = inputWorkbook.GetSheet("Page1");


        // Create a new worksheet in the output workbook
        ISheet outputSheet = workbook.GetSheet("Output");

        var size = inputSheet.LastRowNum - 1;
        // Iterate over the rows in the input worksheet
        for (int i = 0; i <= size; i++)
        {
            // Get the current row
            IRow inputRow = inputSheet.GetRow(i+1);

            // Create a new row in the output worksheet
            IRow outputRow = outputSheet.CreateRow(i+1);
            IRow outddputRow = outputSheet.GetRow(i);

            // Get the values of the first and second columns in the input row
            string persionNumber = inputRow.GetCell(0).ToString() ?? "";
            string nameAndFamilly = inputRow.GetCell(1).ToString() ?? "";

            // Create cells in the output row and write the values
            outputRow.CreateCell(0).SetCellValue(persionNumber);
            outputRow.CreateCell(1).SetCellValue(nameAndFamilly);
        }

        return workbook;

    }
}

