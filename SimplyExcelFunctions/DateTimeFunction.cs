namespace SimplyExcelFunctions;

using System;
using System.Globalization;
using ExcelDna.Integration;
using Microsoft.VisualBasic;



public class DateTimeFunction 
{
    [ExcelFunction(
        Description = "reformat to DateTime From String yyyy-mm-dd hh:mm:ss",
        IsVolatile = true,
        Category = "DateTime Functions"
    )]
    public static object reDateTime(string input)
    {
        if (string.IsNullOrEmpty(input) || input.Length < 19)
            return "Invalid Input";

        try
        {
            DateTime dt = DateTime.ParseExact(
                input.Substring(0, 19),
                "yyyy-MM-dd HH:mm:ss",
                CultureInfo.InvariantCulture
            );
            return dt;
        }
        catch (Exception)
        {
            return ExcelError.ExcelErrorValue;
        }
    }

    [ExcelFunction(
        Description = "Convert Unix timestamp to DateTime",
        IsVolatile = true,
        Category = "DateTime Functions"
    )]
    public static object unixToDateTime(string input)
    {
        if (string.IsNullOrEmpty(input))
            return "Invalid Input";

        if (long.TryParse(input, out long unixTimestamp))
        {
            try
            {
                DateTime utcDateTime = DateTimeOffset
                    .FromUnixTimeSeconds(unixTimestamp)
                    .UtcDateTime;
                DateTime localDateTime = utcDateTime.ToLocalTime();
                return localDateTime;
            }
            catch
            {
                return ExcelError.ExcelErrorValue;
            }
        }
        return ExcelError.ExcelErrorValue;
    }
    
    [ExcelFunction(
        Description = "Convert DateTime to Unix",
        IsVolatile = true,
        Category = "DateTime Functions"
    )]
    public static object dateTimeToUnix(object input)
    {
        DateTime tempDateTime;
        if (input is DateTime dt){
            tempDateTime = dt;
        }
        else if (input is double dbl){
            try 
            {
                tempDateTime = DateTime.FromOADate(dbl);
            }
            catch 
            {
                return ExcelError.ExcelErrorValue;
            }
        }
        else if (input is string str && DateTime.TryParse(str, out tempDateTime)){}        
        else{
            return ExcelError.ExcelErrorValue;
        }
        if (tempDateTime.Kind == DateTimeKind.Unspecified){
            tempDateTime = DateTime.SpecifyKind(tempDateTime, DateTimeKind.Local);   
        }
        try{
            long unixTimestamp = new DateTimeOffset(tempDateTime).ToUnixTimeSeconds();
            return unixTimestamp;
        }
        catch{
            return ExcelError.ExcelErrorValue;
        }
    }
}
