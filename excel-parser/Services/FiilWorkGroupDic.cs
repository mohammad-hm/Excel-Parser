using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

public static class FiilWorkGroupDic
{
    private static readonly Dictionary<string, string> workGroupDic = [];

    public static Dictionary<string, string> Execute(string perNumAndwrGroupfilePath)
    {

        // Create a workbook object from the input file
        IWorkbook inputWorkbook = WorkbookFactory.Create(perNumAndwrGroupfilePath);
        // Get the first worksheet in the input workbook
        ISheet inputSheet = inputWorkbook.GetSheet("sheet");

        var size = inputSheet.LastRowNum - 1;
        // Iterate over the rows in the input worksheet
        for (int i = 0; i <= size; i++)
        {
            // Get the current row
            IRow inputRow = inputSheet.GetRow(i+1);

            string personNumber = inputRow.GetCell(1).ToString() ?? "";
            string workGroup = inputRow.GetCell(4).ToString() ?? "";
            workGroupDic.Add(personNumber, workGroup);
        }
        return workGroupDic;
    }
}

