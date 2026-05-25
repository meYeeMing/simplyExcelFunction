using ExcelDna.Integration;
using SimplyExcelFunctions;
using Xunit;

namespace DateTimeFunction.Tests
{
    public class DateTimeFunctionsTests
    {
        [Theory]
        // Valid: 2026-04-08
        [InlineData("1775628780", 2026, 4, 8, 14, 13, 0)]
        // Valid: The Unix Epoch
        [InlineData("0", 1970, 1, 1, 8, 0, 0)]
        public void UnixToDateTime_ValidInput_ReturnsCorrectDateTime(
            string input,
            int year,
            int month,
            int day,
            int hour,
            int min,
            int sec
        )
        {
            var result = SimplyExcelFunctions.DateTimeFunction.unixToDateTime(input);

            Assert.IsType<DateTime>(result);
            var dt = (DateTime)result;
            Assert.Equal(year, dt.Year);
            Assert.Equal(month, dt.Month);
            Assert.Equal(day, dt.Day);
            Assert.Equal(hour, dt.Hour);
            Assert.Equal(min, dt.Minute);
            Assert.Equal(sec, dt.Second);
        }

        [Fact]
        public void UnixToDateTime_EmptyInput_ReturnsErrorMessage()
        {
            var result = SimplyExcelFunctions.DateTimeFunction.unixToDateTime("");
            Assert.Equal("Invalid Input", result);
        }

        [Theory]
        [InlineData("abc", ExcelError.ExcelErrorValue)]
        [InlineData("-99999999999", ExcelError.ExcelErrorValue)]
        [InlineData("2026-04-08", ExcelError.ExcelErrorValue)]
        public void UnixToDateTime_InvalidInput_ReturnsErrorMessage(string input, object expected)
        {
            var result = SimplyExcelFunctions.DateTimeFunction.unixToDateTime(input);
            Assert.Equal(expected, result);
        }
        [Theory]
        [InlineData("2026-04-08 14:13:00", 1775628780)]
        [InlineData("23/04/2026 14:13:00", 1776924780)]
        [InlineData(46120.60, 1775629440)]
        public void DateTimeToUnix_ValidInput_ReturnsCorrectUnixTimestamp(object input, long expected)
        {

            var result = SimplyExcelFunctions.DateTimeFunction.dateTimeToUnix(input);
            Assert.IsType<long>(result);
            var unixTimestamp = (long)result;
            Assert.Equal(expected, unixTimestamp);
        }
        [Fact]
        public void DateTimeToUnix_DateTimeInput_ReturnsCorrectUnixTimestamp()
        {
            DateTime input = new DateTime(2026, 5, 25, 10, 0, 0);
            long expected = 1779674400;
            var result = SimplyExcelFunctions.DateTimeFunction.dateTimeToUnix(input);
            Assert.Equal(expected, result);
        }
        [Theory]
        [InlineData("2026-13-02",ExcelError.ExcelErrorValue)]
        public void DateTimeToUnix_InvalidInput_ReturnsErrorMessage(object input, object expected)
        {
            var result = SimplyExcelFunctions.DateTimeFunction.dateTimeToUnix(input);
            Assert.Equal(expected, result);
        }
    }
}
