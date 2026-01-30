namespace SimplyExcelFunctions;

using ExcelDna.Integration;
using System;
using System.Globalization;
using Microsoft.VisualBasic;

public class DateTimeFunction
{
    [ExcelFunction(Description = "reformat to DateTime From String yyyy-mm-dd hh:mm:ss", IsVolatile = true, Category = "DateTime Functions")]
    public static object reDateTime(string input)
    {
        if (string.IsNullOrEmpty(input) || input.Length < 19)
            return "Invalid Input";

        try
        {
            DateTime dt = DateTime.ParseExact(input.Substring(0, 19), "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            return dt;
        }
        catch (Exception)
        {
            return ExcelError.ExcelErrorValue;
        }
    }
}
