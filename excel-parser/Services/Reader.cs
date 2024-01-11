using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Text.RegularExpressions;

namespace excel_parser.Services
{
    public class Reader
    {
        public static void ExcelReader(string filePath)
        {
            Console.WriteLine($"Start reading Excel file: {filePath}");

            IWorkbook workbook = WorkbookFactory.Create(filePath);
            ISheet sheet = workbook.GetSheet("PernoPage");

            var persNumAndSumIndexAndworkFeiledDic = new Dictionary<string, Dictionary<string, int>>();
            var singlePersonNumber = string.Empty;
            var singleWorkGroup = string.Empty;
            var startSumIndex = -1;
            var cellIndex = 18;

            for (int i = 0; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    var personNumber = FindPersonNumber(row);
                    if (!string.IsNullOrEmpty(personNumber))
                    {
                        singlePersonNumber = personNumber;
                    }
                    var workGroup = FindWorkGroup(row);
                    if (!string.IsNullOrEmpty(workGroup))
                    {
                        singleWorkGroup = workGroup;
                    }
                    int startingCellIndex = GetColumnIndexByHeader(row, "مجموع");
                    if (startingCellIndex != -1)
                    {
                        startSumIndex = i;
                    }
                    if (startSumIndex != -1 && !string.IsNullOrEmpty(singlePersonNumber) && !string.IsNullOrEmpty(singleWorkGroup))
                    {
                        persNumAndSumIndexAndworkFeiledDic.TryAdd(singlePersonNumber, new Dictionary<string, int>
                        {
                            { singleWorkGroup, startSumIndex }
                        });
                        singlePersonNumber = string.Empty;
                        singleWorkGroup = string.Empty;
                        startSumIndex = -1;
                    }
                }
            }

            var calculateModels = new List<calculateModel>();
            foreach (var kvp in persNumAndSumIndexAndworkFeiledDic)
            {
                var PersonNumber = kvp.Key;

                foreach (var item in kvp.Value)
                {
                    var rowIndex = item.Value;

                    if (rowIndex >= 0 && rowIndex <= sheet.LastRowNum)
                    {
                        IRow row = sheet.GetRow(rowIndex);

                        if (row != null && cellIndex >= 0 && cellIndex < row.LastCellNum)
                        {
                            calculateModel model = new calculateModel
                            {
                                PersonNumber = PersonNumber,
                                WorkGroup = item.Key,
                                TotalFinancialFunction = GetCellValue(row, cellIndex - 1),
                                OvertimeAuthorizedNormal = GetCellValue(row, cellIndex - 2),
                                OvertimeUnauthorizedNormal = GetCellValue(row, cellIndex - 3),
                                OvertimeAuthorizedHolidays = GetCellValue(row, cellIndex - 4),
                                OvertimeUnauthorizedHolidays = GetCellValue(row, cellIndex - 5),
                                OvertimeOnMission = GetCellValue(row, cellIndex - 6),
                                OvertimeSpecial = GetCellValue(row, cellIndex - 7),
                                FractionOfWorkDelay = GetCellValue(row, cellIndex - 8),
                                FractionOfWorkRush = GetCellValue(row, cellIndex - 9),
                                FractionOfWorkUnauthorizedExit = GetCellValue(row, cellIndex - 10),
                                FractionOfWorkAbsence = GetCellValue(row, cellIndex - 11),
                                FractionOfWorkSpecial = GetCellValue(row, cellIndex - 12),
                                WorkingAtNight = GetCellValue(row, cellIndex - 13),
                                AuthorizedDelay = GetCellValue(row, cellIndex - 14),
                                AuthorizedRush = GetCellValue(row, cellIndex - 15),
                                AllMission = GetCellValue(row, cellIndex - 16),
                                UnpaidLeave = GetCellValue(row, cellIndex - 17),
                                PaidLeave = GetCellValue(row, cellIndex - 18),
                            };

                            calculateModels.Add(model);
                        }
                        else
                        {
                            Console.WriteLine($"Invalid Cell Index: {cellIndex}");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"Invalid Row Index: {rowIndex}");
                    }
                }
            }

            foreach (var model in calculateModels)
            {
                switch (model.WorkGroup)
                {
                    case "اداري":
                        //  Console.WriteLine($"Processing {model.WorkGroup} for PersonNumber: {model.PersonNumber}");
                        break;

                    case "پشتيباني":
                        Console.WriteLine($"Processing for PersonNumber: {model.PersonNumber}");
                        SumNumericFields(model);
                        break;

                    case "شناوري 7 تا 8":
                        // Console.WriteLine($"Processing {model.WorkGroup} for PersonNumber: {model.PersonNumber}");
                        break;
                    case "ساعتي":
                        // Console.WriteLine($"Processing {model.WorkGroup} for PersonNumber: {model.PersonNumber}");
                        break;
                    case "داود موسوي":
                        // Console.WriteLine($"Processing {model.WorkGroup} for PersonNumber: {model.PersonNumber}");
                        break;
                    case "فني 8 صبح(بدون شناوري)?":
                        // Console.WriteLine($"Processing {model.WorkGroup} for PersonNumber: {model.PersonNumber}");
                        break;
                    case "فني-شيراز":
                        // Console.WriteLine($"Processing {model.WorkGroup} for PersonNumber: {model.PersonNumber}");
                        break;
                    case "خدمه7تا17":
                        // Console.WriteLine($"Processing {model.WorkGroup} for PersonNumber: {model.PersonNumber}");
                        break;
                    case "شناوري 8-9":
                        //Console.WriteLine($"Processing {model.WorkGroup} for PersonNumber: {model.PersonNumber}");
                        break;
                    case "wsg_2638_27":
                        //   Console.WriteLine($"Processing {model.WorkGroup} for PersonNumber: {model.PersonNumber}");
                        break;
                    case "شناوري 7 تا 10":
                        //   Console.WriteLine($"Processing {model.WorkGroup} for PersonNumber: {model.PersonNumber}");
                        break;
                    default:
                        break;
                }
            }

        }
        public static void WriteToExcel(string filePath, string sheetName)
        {
            Console.WriteLine("WriteToExcel start");

            using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet(sheetName);

                // Specify row and cell for writing data
                int row = 0;
                int cell = 0;

                var newRow = sheet.CreateRow(row++);
                newRow.CreateCell(cell).SetCellValue("IsEnabled");
                newRow.CreateCell(cell + 1).SetCellValue("yes");

                newRow = sheet.CreateRow(row++);
                newRow.CreateCell(cell).SetCellValue("NumberOfActiveClients");
                newRow.CreateCell(cell + 1).SetCellValue("5");

                newRow = sheet.CreateRow(row++);
                newRow.CreateCell(cell).SetCellValue("ExpirationTime");
                newRow.CreateCell(cell + 1).SetCellValue("test");

                workbook.Write(fs);

                Console.WriteLine($"Excel file written to: {filePath}");
            }

        }
        private static string FindPersonNumber(IRow row)
        {
            for (int j = 0; j < row.LastCellNum; j++)
            {
                ICell cellValue = row.GetCell(j);
                if (cellValue != null)
                {
                    if (Regex.IsMatch(cellValue.ToString() ?? "", @"شماره پرسنلی\s*:\s*(\d+)"))
                    {
                        // Extract the number using regular expression
                        Match match = Regex.Match(cellValue.ToString() ?? "", @"شماره پرسنلی\s*:\s*(\d+)");
                        if (match.Success)
                        {
                            var res = match.Groups[1].Value;
                            return (res);
                        }
                    }

                }
            }

            return "";
        }

        private static string FindWorkGroup(IRow row)
        {
            const string headerPattern = @"گروه کاری\s*:\s*(.*)";

            for (int j = 0; j < row.LastCellNum; j++)
            {
                ICell cell = row.GetCell(j);

                if (cell != null)
                {
                    string cellValue = cell.ToString() ?? "";

                    // Use regular expression to match the header pattern
                    Match match = Regex.Match(cellValue ?? "", headerPattern);
                    if (match.Success)
                    {
                        // Extract the value from the matched group
                        return match.Groups[1].Value.Trim();
                    }
                }
            }

            return "";
        }

        private static int GetColumnIndexByHeader(IRow headerRow, string header)
        {
            for (int i = 0; i < headerRow.LastCellNum; i++)
            {
                if (headerRow.GetCell(i)?.ToString()?.Trim() == header)
                {
                    return i;
                }
            }
            return -1;
        }

        private static string GetCellValue(IRow row, int cellIndex)
        {
            ICell cell = row.GetCell(cellIndex);
            return cell?.ToString() ?? string.Empty;
        }

        private static void SumNumericFields(calculateModel model)
        {
            // Extract all numeric fields and sum them up
            int sum = 0;
            foreach (var property in typeof(calculateModel).GetProperties())
            {
                if (property.PropertyType == typeof(string) && property.Name != "PersonNumber" && property.Name != "WorkGroup")
                {
                    string value = (string)property.GetValue(model) ?? "0";
                    if (int.TryParse(value, out int numericValue))
                    {
                        sum += numericValue;
                    }
                }
            }

            Console.WriteLine($"The sum of all field is:{sum}");
        }
    }

    public class calculateModel()
    {
        public string PersonNumber { get; set; } = string.Empty;
        public string WorkGroup { get; set; } = string.Empty;
        public string TotalFinancialFunction { get; set; } = string.Empty;
        public string OvertimeAuthorizedNormal { get; set; } = string.Empty;
        public string OvertimeUnauthorizedNormal { get; set; } = string.Empty;
        public string OvertimeAuthorizedHolidays { get; set; } = string.Empty;
        public string OvertimeUnauthorizedHolidays { get; set; } = string.Empty;
        public string OvertimeOnMission { get; set; } = string.Empty;
        public string OvertimeSpecial { get; set; } = string.Empty;
        public string FractionOfWorkDelay { get; set; } = string.Empty;
        public string FractionOfWorkRush { get; set; } = string.Empty;
        public string FractionOfWorkUnauthorizedExit { get; set; } = string.Empty;
        public string FractionOfWorkAbsence { get; set; } = string.Empty;
        public string FractionOfWorkSpecial { get; set; } = string.Empty;
        public string WorkingAtNight { get; set; } = string.Empty;
        public string AuthorizedDelay { get; set; } = string.Empty;
        public string AuthorizedRush { get; set; } = string.Empty;
        public string AllMission { get; set; } = string.Empty;
        public string UnpaidLeave { get; set; } = string.Empty;
        public string PaidLeave { get; set; } = string.Empty;
    }

}
