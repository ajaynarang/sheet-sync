using ClosedXML.Excel;
using System;
using System.IO;
using System.Linq;

class Program
{
    static void Main()
    {
        string masterTemplatePath = "path/to/masterTemplate.xlsx";
        string clientsFolderPath = "path/to/clients"; // Folder containing client data sheets
        string outputPath = "path/to/output"; // Folder to save updated client sheets

        UpdateClientSheets(masterTemplatePath, clientsFolderPath, outputPath);

        Console.WriteLine("Update completed successfully.");
    }

    static void UpdateClientSheets(string masterTemplatePath, string clientsFolderPath, string outputPath)
    {
        // Load master template
        using (var masterWorkbook = new XLWorkbook(masterTemplatePath))
        {
            foreach (var clientFilePath in Directory.GetFiles(clientsFolderPath, "*.xlsx"))
            {
                // Load client data sheet
                using (var clientWorkbook = new XLWorkbook(clientFilePath))
                {
                    var masterSheet = masterWorkbook.Worksheets.First();
                    var clientSheet = clientWorkbook.Worksheets.First();

                    // Update specific columns' data and formatting
                    UpdateSpecificColumns(masterSheet, clientSheet);

                    // Save the updated client sheet to the output folder
                    string clientFileName = Path.GetFileName(clientFilePath);
                    string outputPathForClient = Path.Combine(outputPath, clientFileName);
                    clientWorkbook.SaveAs(outputPathForClient);
                }
            }
        }
    }

    static void UpdateSpecificColumns(IXLWorksheet sourceSheet, IXLWorksheet destinationSheet)
    {
        // Assuming "KeyColumn" is the key column used for matching rows
        string keyColumn = "KeyColumn";

        // List of columns to be updated from client to template
        var columnsToCopy = new[] { "Column1", "Column2", "Column3" }; // Add your column names here

        // Find the column index of the key column
        var keyColumnIndex = sourceSheet.ColumnsUsed().First(col => col.FirstCell().Value.ToString() == keyColumn).ColumnNumber();

        // Iterate through each row in the client sheet
        foreach (var destinationRow in destinationSheet.RowsUsed())
        {
            // Find the corresponding row in the master sheet based on the key column
            var matchingMasterRow = sourceSheet.RowsUsed().FirstOrDefault(row => row.Cell(keyColumnIndex).Value.ToString() == destinationRow.Cell(keyColumn).Value.ToString());

            if (matchingMasterRow != null)
            {
                // Copy data and formatting for specific columns from client to template
                foreach (var columnName in columnsToCopy)
                {
                    var sourceCell = matchingMasterRow.Cell(sourceSheet.ColumnsUsed().First(col => col.FirstCell().Value.ToString() == columnName).ColumnNumber());
                    var destinationCell = destinationRow.Cell(destinationSheet.ColumnsUsed().First(col => col.FirstCell().Value.ToString() == columnName).ColumnNumber());

                    // Update cell value
                    destinationCell.Value = sourceCell.Value;

                    // Copy all formatting properties
                    destinationCell.Style = sourceCell.Style;
                }
            }
        }
    }
}
