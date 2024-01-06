using OfficeOpenXml;
using Spectre.Console;
using System.Globalization;

namespace ExcelReader;

internal class FileHelper
{
    internal static FileInfo GetFile()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var path = AnsiConsole.Ask<string>("Please input the path of the file:");
        var fileName = AnsiConsole.Ask<string>("Please input the name of the file:");

        var filePath = Path.Combine(path, fileName + ".xlsx");
        var file = new FileInfo(filePath);

        return file;
    }

    internal static async Task<List<People>> GetInfoFromFile(FileInfo file)
    {
        var people = new List<People>();

        using var package = new ExcelPackage(file);

        await package.LoadAsync(file);

        var ws = package.Workbook.Worksheets[0];

        int row = 2;
        int col = 1;

        while (string.IsNullOrWhiteSpace(ws.Cells[row, col].Value?.ToString()) == false)
        {
            var birthDate = ws.Cells[row, col + 6].Value.ToString()[..9];

            var person = new People
            {
                Id = int.Parse(ws.Cells[row, col].Value.ToString()),
                FirstName = ws.Cells[row, col + 1].Value.ToString(),
                LastName = ws.Cells[row, col + 2].Value.ToString(),
                Sex = ws.Cells[row, col + 3].Value.ToString(),
                Email = ws.Cells[row, col + 4].Value.ToString(),
                Phone = ws.Cells[row, col + 5].Value.ToString(),
                BirthDate = DateOnly.Parse(birthDate, new CultureInfo("en-US", true)),
                JobTitle = ws.Cells[row, col + 7].Value.ToString()
            };

            people.Add(person);
            row++;
        }

        return people;
    }

    internal static async void AddDataToFile(FileInfo file)
    {
        if (AnsiConsole.Confirm("Would you like to add more data?"))
        {
            var people = GetPeople();

            using var package = new ExcelPackage(file);

            var ws = package.Workbook.Worksheets[0];

            int rowNumber = 1;

            while (!string.IsNullOrWhiteSpace(ws.Cells[$"A{rowNumber}"].Value?.ToString()))
            {
                rowNumber++;
            }

            var range = ws.Cells[$"A{rowNumber}"].LoadFromCollection(people, false);

            range.AutoFitColumns();

            await package.SaveAsync();

            Console.WriteLine("\nThe data has been successfully added.");
        }
    }

    internal static List<People> GetPeople()
    {
        var people = new List<People>();
        bool addMorePeople;
        do
        {
            var person = new People();

            person.Id = AnsiConsole.Ask<int>("Person's id:");
            person.FirstName = AnsiConsole.Ask<string>("Person's first name:");
            person.LastName = AnsiConsole.Ask<string>("Person's last name:");
            person.Sex = AnsiConsole.Ask<string>("Person's sex:");
            person.Email = AnsiConsole.Ask<string>("Person's email:");
            person.Phone = AnsiConsole.Ask<string>("Person's phone:");
            person.BirthDate = AnsiConsole.Ask<DateOnly>("Person's birth date (MM-dd-yyyy):");
            person.JobTitle = AnsiConsole.Ask<string>("Person's job title:");

            people.Add(person);

            addMorePeople = AnsiConsole.Confirm("Would you like to add more people?");
        } while (addMorePeople);

        return people;
    }
}
