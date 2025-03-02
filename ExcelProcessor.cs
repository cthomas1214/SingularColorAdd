using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

class ExcelProcessor
{
    private readonly string _folderPath;
    private readonly string _connectionString;

    public ExcelProcessor(string folderPath, string connectionString)
    {
        _folderPath = folderPath;
        _connectionString = connectionString;
    }

    //Goes through each file in folder path
    public void ProcessFiles()
    {
        foreach (string filePath in Directory.EnumerateFiles(_folderPath, "*.xls*"))
        {
            Console.WriteLine($"File found: {filePath}");
            using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                XSSFWorkbook workbook = new XSSFWorkbook(fileStream);
                foreach (ISheet sheet in workbook)
                {
                    ProcessSheet(sheet);
                }
            }
        }
    }

    //Processes each Excel file
    private void ProcessSheet(ISheet sheet)
    {
        string nextCellItemName = "";
        string nextCellItemNameNoSpace = "";
        string nextCellColorString = "";
        string nextCellSKUString = "";
        List<string> colorStrings = new List<string>();
        //Goes through each row
        for (int row = 0; row <= sheet.LastRowNum; row++)
        {
            IRow currentRow = sheet.GetRow(row);
            if (currentRow != null)
            {   //Goes through each Column
                for (int col = 0; col < currentRow.LastCellNum; col++)
                {   //Looks for cell called "Item Name" and gets next cell over
                    ICell cell = currentRow.GetCell(col);
                    if (cell != null && cell.CellType == CellType.String && cell.StringCellValue == "Item Name")
                    {
                        ICell nextCell = currentRow.GetCell(col + 1);
                        nextCellItemName = nextCell.StringCellValue;
                        nextCellItemNameNoSpace = nextCellItemName.Trim();
                    }
                    //Searches for all cells with a name called "COLOR NAME"
                    if (cell != null && cell.CellType == CellType.String && cell.StringCellValue == "COLOR NAME" && row != 3)
                    {
                        IRow nextRow = sheet.GetRow(row + 1);
                        ICell nextCellColor = nextRow.GetCell(col);

                        //Gets cells that have a cell of "COLOR NAME" above them
                        if (nextCellColor != null && nextCellColor.CellType == CellType.String)
                        {
                            nextCellColorString = nextCellColor.StringCellValue;
                            if (!string.IsNullOrEmpty(nextCellColorString))
                            {
                                colorStrings.Add(nextCellColorString);
                            }
                        }
                    }
                    //Searched for a cell with a name called "SKU"
                    if (cell != null && cell.CellType == CellType.String && cell.StringCellValue == "SKU" && row < 14 && row != 3)
                    {
                        string skuName = cell.StringCellValue;

                        //Gets the next column of found word of "SKU"
                        ICell nextCellSKU = currentRow.GetCell(col + 1);
                        if (nextCellSKU.CellType == CellType.String)
                        {
                            nextCellSKUString = nextCellSKU.StringCellValue;
                        }
                        //If cell is number, cell can be a number
                        else if (nextCellSKU.CellType == CellType.Numeric)
                        {
                            double nextCellSKUInt = nextCellSKU.NumericCellValue;
                            nextCellSKUString = nextCellSKUInt.ToString();
                        }
                    }
                }//for (int col = 0; col < currentRow.LastCellNum; col++)
            }//if (currentRow != null)
        }//for (int row = 0; row <= sheet.LastRowNum; row++)

        HandleColors(nextCellItemName, nextCellItemNameNoSpace, nextCellSKUString, colorStrings);
    }//private void ProcessSheet(ISheet sheet)


    //Looks for single color and logs product name and sku to help query for missed products with inconsistencies
    private void HandleColors(string nextCellItemName, string nextCellItemNameNoSpace, string nextCellSKUString, List<string> colorStrings)
    {
        string singleColorString;
        string singleColorAndText;

        //Multiple colors
        if (colorStrings.Count > 1)
        {
            Console.WriteLine("Multiple colors:");
            Console.WriteLine($"NAME: {nextCellItemName}");
            Console.WriteLine($"SKU: {nextCellSKUString}");
            foreach (string color in colorStrings)
            {
                Console.WriteLine("-------------Multiple colors: " + color);
            }
        }
        //Single color being what we need
        else if (colorStrings.Count == 1)
        {
            singleColorString = colorStrings[0];
            Console.WriteLine($"NAME: {nextCellItemName}");
            Console.WriteLine($"SKU:{nextCellSKUString}");
            Console.WriteLine("Single color: " + singleColorString);

            //Format to add the color in HTML
            singleColorAndText = $"<p>Color: {singleColorString}</p>";
            try
            {
                UpdateDataInSqlServer(singleColorAndText, nextCellItemName);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ERROR: {ex.Message}");
                Console.WriteLine($"failed to update: {nextCellItemName} {singleColorAndText}");
                Console.WriteLine($"-------------FAILED NAME: {nextCellItemName}");
                Console.WriteLine($"-------------FAILED SKU:{nextCellSKUString}");
                UpdateDataInSqlServer(singleColorAndText, nextCellItemNameNoSpace);
            }
        }
        //No Colors
        else
        {
            Console.WriteLine("No colors found.");
            Console.WriteLine($"-------------------NO COLORS FOUND NAME: {nextCellItemName}");
            Console.WriteLine($"-------------------NO COLORS FOUND SKU:{nextCellSKUString}");
        }
    }//private void HandleColors()


    //Runs query to add associated color to product description of products with one color
    private void UpdateDataInSqlServer(string singleColorAndText, string nextCellItemName)
    {
        using (SqlConnection connection = new SqlConnection(_connectionString))
        {
            connection.Open();
            string query = "UPDATE Products SET ProductDescription = @Color + COALESCE(ProductDescription, '') " +
                           "WHERE ProductName = @ProdName AND ProductDescription NOT LIKE '%' + 'Color:' + '%'";
            using (SqlCommand command = new SqlCommand(query, connection))
            {
                command.Parameters.AddWithValue("@Color", singleColorAndText);
                command.Parameters.AddWithValue("@ProdName", nextCellItemName);
                command.ExecuteNonQuery();
            }
        }
    }//private void UpdateDataInSqlServer()
}