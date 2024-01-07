using excel_parser.Services;
using Microsoft.Extensions.Hosting;

var host = Host.CreateDefaultBuilder().ConfigureServices((hostContext, services) =>
{
    Console.WriteLine("Hello my dear");
    Console.WriteLine("Please enter the path to the Excel file:");
    string filePath = Console.ReadLine() ?? "";

    // Validate if the file exists, etc., before proceeding
    if (!System.IO.File.Exists(filePath))
    {
        Console.WriteLine("Invalid file path or file does not exist.");
        return;
    }
    Reader.ExcelReader(filePath);
 
}).Build();

host.Run();
