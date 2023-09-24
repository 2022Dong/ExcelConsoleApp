using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Drawing;

// 24/9/2023
namespace ExcelConsole
{
    partial class Program
    {
        static async Task Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // create file location
            var file = new FileInfo(@"C:\Users\Dongyun Huang\Source\Repos\ExcelConsoleApp\Demo.xlsx");

            var people = GetSetupData();

            // async save
            await SaveExcelFile(people, file);

            // read
            List<PersonModel> peopleFromExcel = await LoadExcelFile(file);

            foreach (var p in peopleFromExcel)
            {
                Console.WriteLine($"{p.Id} {p.FirstName} {p.LastName}");
            }
        }

        private static async Task<List<PersonModel>> LoadExcelFile(FileInfo file)
        {
            List<PersonModel> output = new();

            using var package = new ExcelPackage(file);

            // need to check if the file exists, before loading
            if (file.Exists)
            {
                await package.LoadAsync(file);

                var ws = package.Workbook.Worksheets[0];

                int row = 3;
                int col = 1;

                while (string.IsNullOrWhiteSpace(ws.Cells[row,col].Value?.ToString()) == false)
                {
                    PersonModel p = new();
                    p.Id = int.Parse(ws.Cells[row, col].Value.ToString());
                    p.FirstName = ws.Cells[row, col + 1].Value.ToString();
                    p.LastName = ws.Cells[row, col + 2].Value.ToString();
                    output.Add(p);
                    row += 1; // if missing -> infinite loop!
                }                
            }
            return output;
        }

        private static async Task SaveExcelFile(List<PersonModel> people, FileInfo file)
        {
            DeleteIfExists(file);

            using var package = new ExcelPackage(file);

            var ws = package.Workbook.Worksheets.Add("MainReport");

            var range = ws.Cells["A2"].LoadFromCollection(people, true);
            range.AutoFitColumns();

            // Formats the header  -- styling
            ws.Cells["A1"].Value = "Our Cool Report";
            ws.Cells["A1:C1"].Merge = true;
            ws.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Row(1).Style.Font.Size = 24;
            ws.Row(1).Style.Font.Color.SetColor(Color.Blue);

            ws.Row(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Row(2).Style.Font.Bold = true;
            ws.Column(3).Width = 20; // override AutoFitColumns()

            await package.SaveAsync();            
        }

        private static void DeleteIfExists(FileInfo file)
        {
            if(file.Exists)
            {
                file.Delete();
            }
        }

        private static List<PersonModel> GetSetupData()
        {
            List<PersonModel> output = new()
            {
                new() {Id = 1, FirstName = "Tim", LastName = "Corey"},
                new() {Id = 2, FirstName = "Sue", LastName = "Storm"},
                new() {Id = 3, FirstName = "Tim", LastName = "Smith"}
            };
            return output;
        }
    }
}
