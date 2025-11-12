// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for business day calculation functions (NETWORKDAYS, WORKDAY, WEEKNUM).
/// </summary>
public class BusinessDayFunctionTests
{
    #region NETWORKDAYS Tests

    [Fact]
    public void Networkdays_ValidDates_NoHolidays_ReturnsWorkingDays()
    {
        var func = NetworkdaysFunction.Instance;

        // NETWORKDAYS(DATE(2024,1,1), DATE(2024,1,31))
        // Jan 2024: 1st is Monday, 31st is Wednesday
        // Should be 23 working days
        var startDate = new DateTime(2024, 1, 1).ToOADate();
        var endDate = new DateTime(2024, 1, 31).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(endDate),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(23.0, result.NumericValue);
    }

    [Fact]
    public void Networkdays_OneWeek_ReturnsCorrectCount()
    {
        var func = NetworkdaysFunction.Instance;

        // Monday to Friday (5 working days)
        var monday = new DateTime(2024, 1, 1).ToOADate(); // Monday
        var friday = new DateTime(2024, 1, 5).ToOADate(); // Friday

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(monday),
            CellValue.FromNumber(friday),
        });

        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void Networkdays_IncludesWeekend_ExcludesSaturdaySunday()
    {
        var func = NetworkdaysFunction.Instance;

        // Monday to Monday (should be 6 working days, skipping Sat/Sun)
        var monday1 = new DateTime(2024, 1, 1).ToOADate(); // Monday
        var monday2 = new DateTime(2024, 1, 8).ToOADate(); // Next Monday

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(monday1),
            CellValue.FromNumber(monday2),
        });

        Assert.Equal(6.0, result.NumericValue);
    }

    [Fact]
    public void Networkdays_ReversedDates_ReturnsNegativeCount()
    {
        var func = NetworkdaysFunction.Instance;

        var laterDate = new DateTime(2024, 1, 31).ToOADate();
        var earlierDate = new DateTime(2024, 1, 1).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(laterDate),
            CellValue.FromNumber(earlierDate),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(-23.0, result.NumericValue);
    }

    [Fact]
    public void Networkdays_SameDate_ReturnsOne()
    {
        var func = NetworkdaysFunction.Instance;

        var date = new DateTime(2024, 1, 2).ToOADate(); // Tuesday

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(date),
            CellValue.FromNumber(date),
        });

        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void Networkdays_WeekendDate_ReturnsZero()
    {
        var func = NetworkdaysFunction.Instance;

        var saturday = new DateTime(2024, 1, 6).ToOADate(); // Saturday

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(saturday),
            CellValue.FromNumber(saturday),
        });

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void Networkdays_WithSingleHoliday_ExcludesHoliday()
    {
        var func = NetworkdaysFunction.Instance;

        // Jan 1-5, 2024 (Mon-Fri = 5 days, but Jan 1 is a holiday)
        var monday = new DateTime(2024, 1, 1).ToOADate();
        var friday = new DateTime(2024, 1, 5).ToOADate();
        var holiday = new DateTime(2024, 1, 1).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(monday),
            CellValue.FromNumber(friday),
            CellValue.FromNumber(holiday),
        });

        Assert.Equal(4.0, result.NumericValue); // 5 days - 1 holiday
    }

    [Fact]
    public void Networkdays_InvalidArguments_ReturnsError()
    {
        var func = NetworkdaysFunction.Instance;

        // Wrong number of arguments
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(44927),
        });

        Assert.True(result1.IsError);
        Assert.Equal("#VALUE!", result1.ErrorValue);

        // Non-numeric argument
        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromString("text"),
            CellValue.FromNumber(44927),
        });

        Assert.True(result2.IsError);
        Assert.Equal("#VALUE!", result2.ErrorValue);
    }

    [Fact]
    public void Networkdays_ErrorPropagation_ReturnsError()
    {
        var func = NetworkdaysFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(44927),
        });

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    #endregion

    #region WORKDAY Tests

    [Fact]
    public void Workday_AddPositiveDays_ReturnsCorrectDate()
    {
        var func = WorkdayFunction.Instance;

        // WORKDAY(DATE(2024,1,1), 10)
        // Start: Jan 1, 2024 (Monday)
        // Add 10 working days: should be Jan 15 (Monday)
        var startDate = new DateTime(2024, 1, 1).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(10),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        var expectedDate = new DateTime(2024, 1, 15).ToOADate();
        Assert.Equal(expectedDate, result.NumericValue);
    }

    [Fact]
    public void Workday_AddZeroDays_ReturnsSameDate()
    {
        var func = WorkdayFunction.Instance;

        var startDate = new DateTime(2024, 1, 2).ToOADate(); // Tuesday

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(0),
        });

        Assert.Equal(startDate, result.NumericValue);
    }

    [Fact]
    public void Workday_AddNegativeDays_ReturnsEarlierDate()
    {
        var func = WorkdayFunction.Instance;

        // WORKDAY(DATE(2024,1,15), -10)
        // Start: Jan 15, 2024 (Monday)
        // Subtract 10 working days: should be Jan 1 (Monday)
        var startDate = new DateTime(2024, 1, 15).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(-10),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        var expectedDate = new DateTime(2024, 1, 1).ToOADate();
        Assert.Equal(expectedDate, result.NumericValue);
    }

    [Fact]
    public void Workday_SkipsWeekends_Correctly()
    {
        var func = WorkdayFunction.Instance;

        // Start on Friday, add 1 day, should be next Monday
        var friday = new DateTime(2024, 1, 5).ToOADate(); // Friday

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(friday),
            CellValue.FromNumber(1),
        });

        var expectedMonday = new DateTime(2024, 1, 8).ToOADate();
        Assert.Equal(expectedMonday, result.NumericValue);
    }

    [Fact]
    public void Workday_WithSingleHoliday_SkipsHoliday()
    {
        var func = WorkdayFunction.Instance;

        // Start Jan 1, add 5 days with Jan 2 as holiday
        // Should skip Jan 2 and land on Jan 8
        var startDate = new DateTime(2024, 1, 1).ToOADate();
        var holiday = new DateTime(2024, 1, 2).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(5),
            CellValue.FromNumber(holiday),
        });

        var expectedDate = new DateTime(2024, 1, 8).ToOADate();
        Assert.Equal(expectedDate, result.NumericValue);
    }

    [Fact]
    public void Workday_InvalidArguments_ReturnsError()
    {
        var func = WorkdayFunction.Instance;

        // Wrong number of arguments
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(44927),
        });

        Assert.True(result1.IsError);
        Assert.Equal("#VALUE!", result1.ErrorValue);

        // Non-numeric argument
        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromString("text"),
            CellValue.FromNumber(10),
        });

        Assert.True(result2.IsError);
        Assert.Equal("#VALUE!", result2.ErrorValue);
    }

    [Fact]
    public void Workday_ErrorPropagation_ReturnsError()
    {
        var func = WorkdayFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(44927),
            CellValue.Error("#REF!"),
        });

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    #endregion

    #region WEEKNUM Tests

    [Fact]
    public void Weeknum_StartOfYear_ReturnsOne()
    {
        var func = WeeknumFunction.Instance;

        // WEEKNUM(DATE(2024,1,1))
        var date = new DateTime(2024, 1, 1).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(date),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void Weeknum_MidYear_ReturnsCorrectWeek()
    {
        var func = WeeknumFunction.Instance;

        // WEEKNUM(DATE(2024,7,1)) - mid-year
        var date = new DateTime(2024, 7, 1).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(date),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue >= 26 && result.NumericValue <= 28);
    }

    [Fact]
    public void Weeknum_Type1_SundayStart_ReturnsCorrectWeek()
    {
        var func = WeeknumFunction.Instance;

        // WEEKNUM(DATE(2024,1,15), 1) - week starts Sunday
        var date = new DateTime(2024, 1, 15).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(date),
            CellValue.FromNumber(1),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue >= 2 && result.NumericValue <= 4);
    }

    [Fact]
    public void Weeknum_Type2_MondayStart_ReturnsCorrectWeek()
    {
        var func = WeeknumFunction.Instance;

        // WEEKNUM(DATE(2024,1,15), 2) - week starts Monday
        var date = new DateTime(2024, 1, 15).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(date),
            CellValue.FromNumber(2),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue >= 2 && result.NumericValue <= 4);
    }

    [Fact]
    public void Weeknum_Type11_ISO8601_ReturnsCorrectWeek()
    {
        var func = WeeknumFunction.Instance;

        // WEEKNUM(DATE(2024,1,15), 11) - ISO 8601
        var date = new DateTime(2024, 1, 15).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(date),
            CellValue.FromNumber(11),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue >= 2 && result.NumericValue <= 4);
    }

    [Fact]
    public void Weeknum_Type21_ISO8601_SameAsType11()
    {
        var func = WeeknumFunction.Instance;

        // WEEKNUM(DATE(2024,1,15), 21) - ISO 8601 (same as 11)
        var date = new DateTime(2024, 1, 15).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(date),
            CellValue.FromNumber(21),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue >= 2 && result.NumericValue <= 4);
    }

    [Fact]
    public void Weeknum_InvalidReturnType_ReturnsError()
    {
        var func = WeeknumFunction.Instance;

        // Invalid return_type (outside 1-21 range)
        var date = new DateTime(2024, 1, 15).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(date),
            CellValue.FromNumber(25),
        });

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Weeknum_InvalidArguments_ReturnsError()
    {
        var func = WeeknumFunction.Instance;

        // Wrong number of arguments
        var result1 = func.Execute(null!, new CellValue[] { });

        Assert.True(result1.IsError);
        Assert.Equal("#VALUE!", result1.ErrorValue);

        // Non-numeric date argument
        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromString("text"),
        });

        Assert.True(result2.IsError);
        Assert.Equal("#VALUE!", result2.ErrorValue);
    }

    [Fact]
    public void Weeknum_ErrorPropagation_ReturnsError()
    {
        var func = WeeknumFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.Error("#N/A"),
        });

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    [Fact]
    public void Weeknum_EndOfYear_ReturnsCorrectWeek()
    {
        var func = WeeknumFunction.Instance;

        // WEEKNUM(DATE(2024,12,31))
        var date = new DateTime(2024, 12, 31).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(date),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue >= 52 && result.NumericValue <= 54);
    }

    #endregion
}
