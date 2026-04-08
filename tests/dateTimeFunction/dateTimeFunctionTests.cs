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
    }
}
