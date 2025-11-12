// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for international Date/Time functions (ISOWEEKNUM, WORKDAY.INTL, NETWORKDAYS.INTL).
/// </summary>
public class DateTimeIntlFunctionTests
{
    #region ISOWEEKNUM Tests

    [Fact]
    public void Isoweeknum_January1_2024_ReturnsCorrectWeek()
    {
        var func = IsoweeknumFunction.Instance;

        // January 1, 2024 is a Monday, which is week 1 in ISO 8601
        var date = new DateTime(2024, 1, 1).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(date),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void Isoweeknum_December31_2023_ReturnsWeek52()
    {
        var func = IsoweeknumFunction.Instance;

        // December 31, 2023 is a Sunday, part of week 52 of 2023
        var date = new DateTime(2023, 12, 31).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(date),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(52.0, result.NumericValue);
    }

    [Fact]
    public void Isoweeknum_MidYear_ReturnsCorrectWeek()
    {
        var func = IsoweeknumFunction.Instance;

        // July 15, 2024 should be around week 29
        var date = new DateTime(2024, 7, 15).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(date),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue >= 28 && result.NumericValue <= 30);
    }

    [Fact]
    public void Isoweeknum_InvalidArguments_ReturnsError()
    {
        var func = IsoweeknumFunction.Instance;

        // Wrong number of arguments
        var result1 = func.Execute(null!, Array.Empty<CellValue>());
        Assert.True(result1.IsError);
        Assert.Equal("#VALUE!", result1.ErrorValue);

        // Non-numeric argument
        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromString("text"),
        });
        Assert.True(result2.IsError);
        Assert.Equal("#VALUE!", result2.ErrorValue);
    }

    [Fact]
    public void Isoweeknum_ErrorPropagation_ReturnsError()
    {
        var func = IsoweeknumFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.Error("#DIV/0!"),
        });

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    #endregion

    #region WORKDAY.INTL Tests

    [Fact]
    public void WorkdayIntl_DefaultWeekend_SkipsSaturdaySunday()
    {
        var func = WorkdayIntlFunction.Instance;

        // Start: Friday Jan 5, 2024, add 1 working day = Monday Jan 8, 2024
        var startDate = new DateTime(2024, 1, 5).ToOADate(); // Friday
        var expectedDate = new DateTime(2024, 1, 8).ToOADate(); // Monday

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(1),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(expectedDate, result.NumericValue);
    }

    [Fact]
    public void WorkdayIntl_SundayMondayWeekend_SkipsSundayMonday()
    {
        var func = WorkdayIntlFunction.Instance;

        // Start: Friday Jan 5, 2024, add 1 working day with weekend=2 (Sun-Mon) = Tuesday Jan 9, 2024
        var startDate = new DateTime(2024, 1, 5).ToOADate(); // Friday
        var expectedDate = new DateTime(2024, 1, 9).ToOADate(); // Tuesday

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(1),
            CellValue.FromNumber(2), // Weekend type 2 = Sunday-Monday
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(expectedDate, result.NumericValue);
    }

    [Fact]
    public void WorkdayIntl_SundayOnlyWeekend_SkipsOnlySunday()
    {
        var func = WorkdayIntlFunction.Instance;

        // Start: Friday Jan 5, 2024, add 1 working day with weekend=11 (Sunday only) = Saturday Jan 6, 2024
        var startDate = new DateTime(2024, 1, 5).ToOADate(); // Friday
        var expectedDate = new DateTime(2024, 1, 6).ToOADate(); // Saturday

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(1),
            CellValue.FromNumber(11), // Weekend type 11 = Sunday only
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(expectedDate, result.NumericValue);
    }

    [Fact]
    public void WorkdayIntl_CustomWeekendString_WorksCorrectly()
    {
        var func = WorkdayIntlFunction.Instance;

        // Start: Monday Jan 1, 2024, add 1 working day with custom weekend "1000000" (Sunday only)
        var startDate = new DateTime(2024, 1, 1).ToOADate(); // Monday
        var expectedDate = new DateTime(2024, 1, 2).ToOADate(); // Tuesday

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(1),
            CellValue.FromString("1000000"), // Sunday only
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(expectedDate, result.NumericValue);
    }

    [Fact]
    public void WorkdayIntl_NegativeDays_GoesBackward()
    {
        var func = WorkdayIntlFunction.Instance;

        // Start: Monday Jan 8, 2024, subtract 1 working day = Friday Jan 5, 2024
        var startDate = new DateTime(2024, 1, 8).ToOADate(); // Monday
        var expectedDate = new DateTime(2024, 1, 5).ToOADate(); // Friday

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(-1),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(expectedDate, result.NumericValue);
    }

    [Fact]
    public void WorkdayIntl_WithHoliday_SkipsHoliday()
    {
        var func = WorkdayIntlFunction.Instance;

        // Start: Friday Jan 5, 2024, add 1 working day, but Jan 8 is a holiday = Tuesday Jan 9, 2024
        var startDate = new DateTime(2024, 1, 5).ToOADate(); // Friday
        var holiday = new DateTime(2024, 1, 8).ToOADate(); // Monday (holiday)
        var expectedDate = new DateTime(2024, 1, 9).ToOADate(); // Tuesday

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(1),
            CellValue.FromNumber(1), // Default weekend
            CellValue.FromNumber(holiday),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(expectedDate, result.NumericValue);
    }

    [Fact]
    public void WorkdayIntl_InvalidWeekendType_ReturnsError()
    {
        var func = WorkdayIntlFunction.Instance;

        var startDate = new DateTime(2024, 1, 1).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(1),
            CellValue.FromNumber(18), // Invalid weekend type
        });

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void WorkdayIntl_InvalidWeekendString_ReturnsError()
    {
        var func = WorkdayIntlFunction.Instance;

        var startDate = new DateTime(2024, 1, 1).ToOADate();

        // Wrong length
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(1),
            CellValue.FromString("10000"), // Too short
        });
        Assert.True(result1.IsError);
        Assert.Equal("#VALUE!", result1.ErrorValue);

        // Invalid characters
        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(1),
            CellValue.FromString("1234567"), // Invalid chars
        });
        Assert.True(result2.IsError);
        Assert.Equal("#VALUE!", result2.ErrorValue);
    }

    #endregion

    #region NETWORKDAYS.INTL Tests

    [Fact]
    public void NetworkdaysIntl_DefaultWeekend_CountsWorkdays()
    {
        var func = NetworkdaysIntlFunction.Instance;

        // From Monday Jan 1, 2024 to Friday Jan 5, 2024 = 5 working days
        var startDate = new DateTime(2024, 1, 1).ToOADate(); // Monday
        var endDate = new DateTime(2024, 1, 5).ToOADate(); // Friday

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(endDate),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void NetworkdaysIntl_IncludesWeekend_CountsMoreDays()
    {
        var func = NetworkdaysIntlFunction.Instance;

        // From Monday Jan 1, 2024 to Sunday Jan 7, 2024 (includes weekend)
        // With default weekend (Sat-Sun), only count Mon-Fri = 5 days
        var startDate = new DateTime(2024, 1, 1).ToOADate(); // Monday
        var endDate = new DateTime(2024, 1, 7).ToOADate(); // Sunday

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(endDate),
        });

        Assert.Equal(5.0, result.NumericValue);
    }

    [Fact]
    public void NetworkdaysIntl_SundayOnlyWeekend_CountsSaturdays()
    {
        var func = NetworkdaysIntlFunction.Instance;

        // From Monday Jan 1, 2024 to Sunday Jan 7, 2024
        // With weekend=11 (Sunday only), count Mon-Sat = 6 days
        var startDate = new DateTime(2024, 1, 1).ToOADate(); // Monday
        var endDate = new DateTime(2024, 1, 7).ToOADate(); // Sunday

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(endDate),
            CellValue.FromNumber(11), // Weekend type 11 = Sunday only
        });

        Assert.Equal(6.0, result.NumericValue);
    }

    [Fact]
    public void NetworkdaysIntl_CustomWeekendString_WorksCorrectly()
    {
        var func = NetworkdaysIntlFunction.Instance;

        // From Monday Jan 1, 2024 to Friday Jan 5, 2024
        // With custom weekend "0000011" (Fri-Sat), only Mon-Thu are workdays = 4 days
        var startDate = new DateTime(2024, 1, 1).ToOADate(); // Monday
        var endDate = new DateTime(2024, 1, 5).ToOADate(); // Friday

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(endDate),
            CellValue.FromString("0000011"), // Friday-Saturday weekend
        });

        Assert.Equal(4.0, result.NumericValue);
    }

    [Fact]
    public void NetworkdaysIntl_ReversedDates_ReturnsNegativeCount()
    {
        var func = NetworkdaysIntlFunction.Instance;

        // From Friday Jan 5, 2024 to Monday Jan 1, 2024 (reversed) = -5 working days
        var startDate = new DateTime(2024, 1, 5).ToOADate(); // Friday
        var endDate = new DateTime(2024, 1, 1).ToOADate(); // Monday

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(endDate),
        });

        Assert.Equal(-5.0, result.NumericValue);
    }

    [Fact]
    public void NetworkdaysIntl_WithHoliday_ExcludesHoliday()
    {
        var func = NetworkdaysIntlFunction.Instance;

        // From Monday Jan 1, 2024 to Friday Jan 5, 2024 with Jan 3 as holiday = 4 working days
        var startDate = new DateTime(2024, 1, 1).ToOADate(); // Monday
        var endDate = new DateTime(2024, 1, 5).ToOADate(); // Friday
        var holiday = new DateTime(2024, 1, 3).ToOADate(); // Wednesday (holiday)

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(endDate),
            CellValue.FromNumber(1), // Default weekend
            CellValue.FromNumber(holiday),
        });

        Assert.Equal(4.0, result.NumericValue);
    }

    [Fact]
    public void NetworkdaysIntl_SameDates_ReturnsOne()
    {
        var func = NetworkdaysIntlFunction.Instance;

        // Same date (Monday) = 1 working day
        var date = new DateTime(2024, 1, 1).ToOADate(); // Monday

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(date),
            CellValue.FromNumber(date),
        });

        Assert.Equal(1.0, result.NumericValue);
    }

    [Fact]
    public void NetworkdaysIntl_SameDatesWeekend_ReturnsZero()
    {
        var func = NetworkdaysIntlFunction.Instance;

        // Same date (Saturday) with default weekend = 0 working days
        var date = new DateTime(2024, 1, 6).ToOADate(); // Saturday

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(date),
            CellValue.FromNumber(date),
        });

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void NetworkdaysIntl_InvalidWeekendType_ReturnsError()
    {
        var func = NetworkdaysIntlFunction.Instance;

        var startDate = new DateTime(2024, 1, 1).ToOADate();
        var endDate = new DateTime(2024, 1, 5).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(endDate),
            CellValue.FromNumber(0), // Invalid weekend type
        });

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void NetworkdaysIntl_InvalidArguments_ReturnsError()
    {
        var func = NetworkdaysIntlFunction.Instance;

        // Wrong number of arguments
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(44927),
        });
        Assert.True(result1.IsError);
        Assert.Equal("#VALUE!", result1.ErrorValue);

        // Non-numeric date argument
        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromString("text"),
            CellValue.FromNumber(44927),
        });
        Assert.True(result2.IsError);
        Assert.Equal("#VALUE!", result2.ErrorValue);
    }

    [Fact]
    public void NetworkdaysIntl_ErrorPropagation_ReturnsError()
    {
        var func = NetworkdaysIntlFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(44927),
            CellValue.Error("#REF!"),
        });

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    #endregion
}
