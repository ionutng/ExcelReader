using OfficeOpenXml;
using System.Globalization;

namespace ExcelReader;

internal class FileHelper
{
    internal FileInfo GetFile()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var path = Path.GetDirectoryName(Environment.CurrentDirectory);
        var filePath = Path.Combine(path, "People.xlsx");
        var file = new FileInfo(filePath);

        return file;
    }

    internal async Task<List<People>> GetInfoFromFile(FileInfo file)
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
}
