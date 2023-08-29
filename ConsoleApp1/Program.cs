using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

class Program
{
    static void Main(string[] args)
    {
        string swaggerJsonFilePath = "Provide the path for Swagger File";
        string excelFilePath = "Provide the path to save the excel with the required name";

        // Loads the Swagger JSON file
        string swaggerJson = File.ReadAllText(swaggerJsonFilePath);

        // Parsing the Swagger JSON data
        var swaggerData = Newtonsoft.Json.JsonConvert.DeserializeObject<SwaggerData>(swaggerJson);

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // Create a new Excel package
        using (var excelPackage = new ExcelPackage())
        {
            // Create a new worksheet
            var worksheet = excelPackage.Workbook.Worksheets.Add("Booking API Endpoints");

            // Write the headers
            worksheet.Cells[1, 1].Value = "Controller Name";
            worksheet.Cells[1, 2].Value = "Endpoint Details";
            worksheet.Cells[1, 3].Value = "HTTP Method";
            worksheet.Cells[1, 4].Value = "Summary";
            worksheet.Cells[1, 5].Value = "Description";

            // Write the API endpoint details
            int row = 2;
            foreach (var path in swaggerData.Paths)
            {
                foreach (var method in path.Value)
                {
                    worksheet.Cells[row, 1].Value = method.Value.Tags[0];
                    worksheet.Cells[row, 2].Value = path.Key;
                    worksheet.Cells[row, 3].Value = method.Key.ToUpper();
                    worksheet.Cells[row, 4].Value = method.Value.Summary;
                    worksheet.Cells[row, 5].Value = method.Value.Description;
                    row++;
                }
            }

            // Save the Excel package
            excelPackage.SaveAs(new FileInfo(excelFilePath));
        }

        Console.WriteLine("Excel file created and saved successfully!");
    }
}

class SwaggerData
{
    public Dictionary<string, Dictionary<string, EndpointData>> Paths { get; set; }
}

class EndpointData
{
    public string Summary { get; set; }
    public string Description { get; set; }
    public string[] Tags { get; set; }
}
