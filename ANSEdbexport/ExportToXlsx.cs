using Microsoft.Data.Sqlite;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Serilog;
using System.Drawing;

namespace ANSEdbexport;

public class ExportToXlsx
{
    private static dynamic cfgObj;
    private static SqliteConnection dbConn;
    private static ExcelPackage outTable;
    
    public ExportToXlsx(dynamic _cfgObj, SqliteConnection _dbConn, string _outPath)
    {
        cfgObj = _cfgObj;
        dbConn = _dbConn;

        if (File.Exists(_outPath)) File.Delete(_outPath);
        outTable = new ExcelPackage(new FileInfo(_outPath));
        if (outTable == null) throw new ArgumentException("Error: Output Excel table is null");
    }

    public void Export()
    {
        dbConn.Open();

        string getTablesQuery = "SELECT name FROM sqlite_master WHERE type='table';";
        SqliteCommand getTablesCmd = new SqliteCommand(getTablesQuery, dbConn);
        using (SqliteDataReader tableReader = getTablesCmd.ExecuteReader())
        {
            while (tableReader.Read())
            {
                string tableName = tableReader["name"].ToString();
                if (string.IsNullOrEmpty(tableName)) continue;
                ProcessTableToXlsx(tableName);
            }
        }
    }
    
    /// <summary>
    /// Processes a single db table in multiple rows of content in a .xlsx file, function will be called multiple times
    /// </summary>
    /// <param name="dbTableName">Name of the db table to process</param>
    private static void ProcessTableToXlsx(string dbTableName)
    {
        // get the "types" field from the config
        var typesCfg = cfgObj["types"];
        if (typesCfg == null) throw new ArgumentException("Error: Config file must contain type definitions");
        if (typesCfg[dbTableName] == null)
        {
            Log.Information($"Table {dbTableName, -30} | No configuration file found for this type, skip table");
            return;
        }

        int colOffset = 1;
        if (cfgObj["options"]["col_offset"] != null) colOffset = (int)cfgObj["options"]["col_offset"];
        
        // find the corresponding worksheet for this type
        string outSheetName = GetSheetName(dbTableName);
        Log.Information($"Table {dbTableName, -30} | Inputs written to worksheet {outSheetName}");
        
        // look for sheet with sheetname, if it doesn't exist, create new sheet with sheetname
        if (outTable.Workbook.Worksheets[outSheetName] == null) SetupNewSheet(outSheetName, colOffset);
        ExcelWorksheet outSheet = outTable.Workbook.Worksheets[outSheetName];
        colOffset++;
        
        // write type header for readability
        WriteTypeHeader(dbTableName, outSheet, colOffset);
        
        // Iterate through all rows of db table
        string getDataQuery = $"SELECT * FROM {dbTableName}";
        SqliteCommand getDataCmd = new SqliteCommand(getDataQuery, dbConn);
        using (SqliteDataReader dataReader = getDataCmd.ExecuteReader())
        {
            // Get row number
            int currRow = outSheet.Dimension.Rows + 1;
            outSheet.InsertRow(currRow, 1);
            while (dataReader.Read())
            {
                // Write typename to first cell
                outSheet.Cells[currRow, colOffset].Value = dbTableName.Trim();
                
                // Write the fields to the rest of the row
                WriteTypeDeclaration(dataReader, outSheet, currRow, colOffset);
                
                // Increment and go to next row
                currRow++;
            }
        }
        outTable.Save();
    }
    
    /// <summary>
    /// Writes one row of the type header (highlighted in green) for legibility
    /// </summary>
    /// <param name="tableName">Name of the db table (and also the ANSE model type)</param>
    /// <param name="outSheet">Worksheet that data should be written to</param>
    private static void WriteTypeHeader(string tableName, ExcelWorksheet outSheet, int colOffset)
    {
        // insert two rows, set current row to the bottommost row
        int currRow = outSheet.Dimension.Rows + 2;
        outSheet.InsertRow(currRow - 1, 2);

        // label the current row with KEYWORD: {typename} for readability
        var typeNameCell = outSheet.Cells[currRow, (colOffset)];
        typeNameCell.Value = $"KEYWORD: {tableName}";
        typeNameCell.Style.Font.Bold = true;
        ResizeColumn(colOffset, $"KEYWORD: {tableName}", outSheet);
        
        // get the corresponding type from db
        string getDataQuery = $"SELECT * FROM {tableName}";
        SqliteCommand getDataCmd = new SqliteCommand(getDataQuery, dbConn);
        using (SqliteDataReader dataReader = getDataCmd.ExecuteReader())
        {
            for (int i = 0; i < dataReader.FieldCount; i++)
            {
                int currCol = (i + 1) + colOffset;
                string fieldName = dataReader.GetName(i);
                outSheet.Cells[currRow, currCol].Value = fieldName.Trim();

                // resize column for readability
                ResizeColumn(currCol, fieldName, outSheet);
            }
            // recolor row for readability
            outSheet.Row(currRow).Style.Fill.PatternType = ExcelFillStyle.Solid;
            outSheet.Row(currRow).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#CCFFCC"));
        }
    }

    /// <summary>
    /// Write one row of a single type declaration from the db to the .xlsx sheet
    /// </summary>
    /// <param name="dataReader">Data reader for reading from the database table</param>
    /// <param name="outSheet">Worksheet that the data should be written to</param>
    /// <param name="colOffset">Column offset from which data should be written</param>
    /// <param name="currRow">Current row fir data to be written to</param>
    private static void WriteTypeDeclaration(SqliteDataReader dataReader, ExcelWorksheet outSheet, int currRow, int colOffset)
    {
        Log.Verbose($"\t\tWriting data to Row {currRow, -4}");
        for (int i = 0; i < dataReader.FieldCount; i++)
        {
            if (dataReader.IsDBNull(i)) continue;
            
            // current column number
            int currCol = (i + 1) + colOffset;
            
            // get information about cell value in db
            // string colName = dataReader.GetName(i); CURRENTLY NOT IN USE
            var colVal = dataReader.GetValue(i);
            Type dataType = dataReader.GetFieldType(i);
            
            // get the reference to the cell that we are writing to
            ExcelRange cell = outSheet.Cells[currRow, currCol];
            
            // write value to cell in sheet
            WriteTypeField(cell, dataType, colVal);
            
            // resize cell for readability
            ResizeColumn(currCol, colVal.ToString(), outSheet);
        }
    }
    
    /// <summary>
    /// Takes one cell of value from the database and write it to one cell in the worksheet
    /// </summary>
    /// <param name="cell">Object that allows the function to interact with the cell</param>
    /// <param name="dataType">The datatype of the db value ie: int, double, etc</param>
    /// <param name="fieldVal">Value of the db field</param>
    private static void WriteTypeField(ExcelRange cell, Type dataType, object? fieldVal)
    {
        if (fieldVal == null)
        {
            cell.Value = string.Empty;
            return;
        }

        if (dataType == typeof(int))
        {
            int valInt = 0;
            if (int.TryParse(fieldVal.ToString(), out valInt))
            {
                cell.Value = valInt;
                return;
            }
        }
        else if (dataType == typeof(double))
        {
            double valDub = 0.0;
            if (double.TryParse(fieldVal.ToString(), out valDub))
            {
                cell.Value = valDub;
                return;
            }
        }
        else if (dataType == typeof(bool))
        {
            int valBool = 0;
            if (int.TryParse(fieldVal.ToString(), out valBool))
            {
                cell.Value = (valBool == 0) ? "FALSE" : "TRUE";
                return;
            }
        }
        // if data type is string or something unparsable
        cell.Value = fieldVal.ToString().Trim();

        var val = cell.Value;
        // log info
        Log.Verbose($"\t\t\tData from db table cast to {dataType} and written to cell {@cell}", val, cell);
    }

    /// <summary>
    /// Creates a new sheet and sets it up for further processing
    /// </summary>
    /// <param name="sheetName">Name of the worksheet to create</param>
    private static void SetupNewSheet(string sheetName, int colOffset)
    {
        outTable.Workbook.Worksheets.Add(sheetName);
        var initCell = outTable.Workbook.Worksheets[sheetName].Cells[1, colOffset];
        initCell.Value = "COMMENT";
        
        // recolour comment cell for readability
        initCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
        initCell.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
        
        // put the sheet name in row 1 
        outTable.Workbook.Worksheets[sheetName].Cells[1, 1 + colOffset].Value = sheetName;
        
        // write info to log
        Log.Verbose($"Created sheet: {sheetName}");
    }

    /// <summary>
    /// According to the type declaration, get the appropriate sheet for data to be written to
    /// </summary>
    /// <param name="cfgObj">Object representation of the configuration json</param>
    /// <param name="tableName">Name of the db table (also the name of the ANSE model type)</param>
    /// <returns></returns>
    private static string GetSheetName(string tableName)
    {
        string sheetName = "INPUTS - MISC";
        var groupings = cfgObj["grouping"];
        foreach (var category in groupings)
        {
            foreach (string type in groupings[category.Name])
            {
                if (type == tableName)
                {
                    sheetName = $"INPUTS - {category.Name}";
                    break;
                }
            }
        }
        return sheetName;
    }

    /// <summary>
    /// Resize a column based on the text within the cell
    /// </summary>
    /// <param name="col">Column to resize</param>
    /// <param name="colText">Text within the column</param>
    /// <param name="ws">Worksheet where the column rests</param>
    private static void ResizeColumn(int col, string colText, ExcelWorksheet ws)
    {
        double currWidth = ws.Column(col).Width;
        double newWidth = string.IsNullOrEmpty(colText) ? 0 : colText.Trim().Length * 1.05;
        ws.Column(col).Width = double.Max(currWidth, newWidth);
    }
}