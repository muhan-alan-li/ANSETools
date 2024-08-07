using Microsoft.Data.Sqlite;
using SharedClasses;
using Serilog;

namespace ANSEdbexport;

public class ExportToText
{
    private static dynamic cfgObj;
    private static SqliteConnection dbConn;
    private static string outPath;
    
    public ExportToText(dynamic _cfgObj, SqliteConnection _dbConn, string _outPath)
    {
        cfgObj = _cfgObj;
        dbConn = _dbConn;
        outPath = _outPath;
    }

    public void Export()
    {
        dbConn.Open();
                
        // iterate through all tables in the db 
        string getTablesQuery = "SELECT name FROM sqlite_master WHERE type='table';";
        SqliteCommand getTablesCmd = new SqliteCommand(getTablesQuery, dbConn);
        using (SqliteDataReader tableReader = getTablesCmd.ExecuteReader())
        {
            File.AppendAllText(outPath, $"----CONTENTS GENERATED FROM: {dbConn.DataSource}----\n\n");
            while (tableReader.Read())
            {
                string tableName = tableReader["name"].ToString();
                if (string.IsNullOrEmpty(tableName)) continue;
                ProcessTableToTxt(tableName);
            }
            File.AppendAllText(outPath, $"----END OF FILE----");
        }
    }
    
     /// <summary>
    /// Processes a single db table into text, this function will be called multiple times, once for each table
    /// </summary>
    /// <param name="dbTableName">Name of the db table to process</param>
    /// <exception cref="ArgumentException">Throws an exception if configuration file is null</exception>
    private static void ProcessTableToTxt(string dbTableName)
    {
        // get the "types" field from the config
        var typesCfg = cfgObj["types"];
        if (typesCfg == null) throw new ArgumentException("Error: Config file must contain type definitions");
        if (typesCfg[dbTableName] == null)
        {
            Log.Information($"Table {dbTableName, -30} | No configuration file found for this type, skip table");
            return;
        }

        /*
         * Check the config file and grab the special characters, these include:
         *
         * Comment character: default to '#' ---> # COMMENT
         * Delimiter (head):  default to '<'
         * Delimiter (tail):  default to '>' ---> <TYPE>
         * Splitter:          default to ';' ---> <TYPE>name;1;2;3.0;;;
         */ 
        string[] specChars   = InputChecker.CheckSpecialChars(cfgObj);
        string   commentChar = specChars[0];
        string   delimitHead = specChars[1];
        string   delimitTail = specChars[2];
        string   splitter    = specChars[3];

        File.AppendAllText(outPath, $"\n{commentChar} Generating text from {dbTableName}...\n\n");
        
        string getDataQuery = $"SELECT * FROM {dbTableName}";
        SqliteCommand getDataCmd = new SqliteCommand(getDataQuery, dbConn);
        using (SqliteDataReader dataReader = getDataCmd.ExecuteReader())
        {
            // iterate through every row of the db table
            while (dataReader.Read())
            {
                string outText = $"{delimitHead}{dbTableName.Trim()}{delimitTail}";

                // iterate through all columns of one row in the db table, this gets converted to one line of text
                for (int i = 0; i < dataReader.FieldCount; i++)
                {
                    var colVal = dataReader.GetValue(i);
                    Type colType = dataReader.GetFieldType(i);
                    outText += $"{ParseCol(colVal, colType)}{splitter}";
                }
                File.AppendAllText(outPath, $"{outText}\n");
            }
        }
    }

    /// <summary>
    /// Parses a column within a row (cell) and returns a string representation of its value
    /// </summary>
    /// <param name="val">The object value from db table</param>
    /// <param name="type">The type of the data ie: int, double, etc</param>
    /// <returns>A string representation of the value from the db</returns>
    private static string ParseCol(object? val, Type type)
    {
        string outStr = string.Empty;
        if (val == null) return outStr;
        
        if (type == typeof(int))
        {
            int valInt = 0;
            if (int.TryParse(val.ToString(), out valInt))
                outStr = valInt.ToString();
        }
        else if (type == typeof(double))
        {
            double valDub = 0.0;
            if (double.TryParse(val.ToString(), out valDub))
                outStr = valDub.ToString();
        }
        else if (type == typeof(bool))
        {
            outStr = (val.ToString() == "1") ? "TRUE" : "FALSE";
        }

        if (outStr == string.Empty) outStr = val.ToString().Trim();

        return outStr;
    }
}