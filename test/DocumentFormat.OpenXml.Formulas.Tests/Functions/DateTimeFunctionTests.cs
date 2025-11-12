// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for Date/Time functions.
/// </summary>
public class DateTimeFunctionTests
{
    [Fact]
    public void Days_ValidDates_ReturnsCorrectDifference()
    {
        var func = DaysFunction.Instance;

        // DAYS(DATE(2024,1,31), DATE(2024,1,1)) = 30
        var date1 = new DateTime(2024, 1, 31).ToOADate();
        var date2 = new DateTime(2024, 1, 1).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(date1),
            CellValue.FromNumber(date2),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(30.0, result.NumericValue);
    }

    [Fact]
    public void Days_NegativeDifference_ReturnsNegativeValue()
    {
        var func = DaysFunction.Instance;

        var date1 = new DateTime(2024, 1, 1).ToOADate();
        var date2 = new DateTime(2024, 1, 31).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(date1),
            CellValue.FromNumber(date2),
        });

        Assert.Equal(-30.0, result.NumericValue);
    }

    [Fact]
    public void Days_SameDates_ReturnsZero()
    {
        var func = DaysFunction.Instance;

        var date = new DateTime(2024, 1, 15).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(date),
            CellValue.FromNumber(date),
        });

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void Days_InvalidArguments_ReturnsError()
    {
        var func = DaysFunction.Instance;

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
    public void Time_ValidComponents_ReturnsCorrectFraction()
    {
        var func = TimeFunction.Instance;

        // TIME(12, 30, 0) = 0.520833... (12:30 PM)
        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(12),
            CellValue.FromNumber(30),
            CellValue.FromNumber(0),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.520833333, result.NumericValue, 6);
    }

    [Fact]
    public void Time_Midnight_ReturnsZero()
    {
        var func = TimeFunction.Instance;

        // TIME(0, 0, 0) = 0
        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(0),
            CellValue.FromNumber(0),
            CellValue.FromNumber(0),
        });

        Assert.Equal(0.0, result.NumericValue);
    }

    [Fact]
    public void Time_MaxTime_ReturnsCorrectFraction()
    {
        var func = TimeFunction.Instance;

        // TIME(23, 59, 59)
        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(23),
            CellValue.FromNumber(59),
            CellValue.FromNumber(59),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue < 1.0);
        Assert.True(result.NumericValue > 0.999);
    }

    [Fact]
    public void Time_NegativeValues_ReturnsError()
    {
        var func = TimeFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(-1),
            CellValue.FromNumber(0),
            CellValue.FromNumber(0),
        });

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Time_InvalidArguments_ReturnsError()
    {
        var func = TimeFunction.Instance;

        // Wrong number of arguments
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(12),
            CellValue.FromNumber(30),
        });

        Assert.True(result1.IsError);
        Assert.Equal("#VALUE!", result1.ErrorValue);

        // Non-numeric argument
        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromString("text"),
            CellValue.FromNumber(30),
            CellValue.FromNumber(0),
        });

        Assert.True(result2.IsError);
        Assert.Equal("#VALUE!", result2.ErrorValue);
    }

    [Fact]
    public void TimeValue_ValidTimeString_ReturnsCorrectFraction()
    {
        var func = TimeValueFunction.Instance;

        // TIMEVALUE("12:30:00") = 0.520833...
        var result = func.Execute(null!, new[]
        {
            CellValue.FromString("12:30:00"),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.520833333, result.NumericValue, 6);
    }

    [Fact]
    public void TimeValue_ShortTimeFormat_ParsesCorrectly()
    {
        var func = TimeValueFunction.Instance;

        // TIMEVALUE("3:30 PM")
        var result = func.Execute(null!, new[]
        {
            CellValue.FromString("3:30 PM"),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue > 0.6 && result.NumericValue < 0.7);
    }

    [Fact]
    public void TimeValue_InvalidTimeString_ReturnsError()
    {
        var func = TimeValueFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromString("invalid"),
        });

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void TimeValue_NonTextArgument_ReturnsError()
    {
        var func = TimeValueFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(123),
        });

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void DateValue_ValidDateString_ReturnsSerialNumber()
    {
        var func = DateValueFunction.Instance;

        // DATEVALUE("2024-01-01")
        var result = func.Execute(null!, new[]
        {
            CellValue.FromString("2024-01-01"),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        var expectedDate = new DateTime(2024, 1, 1).ToOADate();
        Assert.Equal(expectedDate, result.NumericValue);
    }

    [Fact]
    public void DateValue_DifferentDateFormat_ParsesCorrectly()
    {
        var func = DateValueFunction.Instance;

        // DATEVALUE("1/15/2024")
        var result = func.Execute(null!, new[]
        {
            CellValue.FromString("1/15/2024"),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        var expectedDate = new DateTime(2024, 1, 15).ToOADate();
        Assert.Equal(expectedDate, result.NumericValue);
    }

    [Fact]
    public void DateValue_InvalidDateString_ReturnsError()
    {
        var func = DateValueFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromString("invalid"),
        });

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void DateValue_NonTextArgument_ReturnsError()
    {
        var func = DateValueFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(123),
        });

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Days360_ValidDates_USMethod_ReturnsCorrectDays()
    {
        var func = Days360Function.Instance;

        // DAYS360(DATE(2024,1,1), DATE(2024,12,31)) = 360
        var startDate = new DateTime(2024, 1, 1).ToOADate();
        var endDate = new DateTime(2024, 12, 31).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(endDate),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(360.0, result.NumericValue);
    }

    [Fact]
    public void Days360_ValidDates_EuropeanMethod_ReturnsCorrectDays()
    {
        var func = Days360Function.Instance;

        // DAYS360(DATE(2024,1,1), DATE(2024,12,31), TRUE)
        var startDate = new DateTime(2024, 1, 1).ToOADate();
        var endDate = new DateTime(2024, 12, 31).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(endDate),
            CellValue.FromBool(true),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(360.0, result.NumericValue);
    }

    [Fact]
    public void Days360_OneMonth_Returns30Days()
    {
        var func = Days360Function.Instance;

        // DAYS360(DATE(2024,1,1), DATE(2024,2,1))
        var startDate = new DateTime(2024, 1, 1).ToOADate();
        var endDate = new DateTime(2024, 2, 1).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(endDate),
        });

        Assert.Equal(30.0, result.NumericValue);
    }

    [Fact]
    public void Days360_InvalidArguments_ReturnsError()
    {
        var func = Days360Function.Instance;

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
    public void Eomonth_ZeroMonths_ReturnsEndOfCurrentMonth()
    {
        var func = EomonthFunction.Instance;

        // EOMONTH(DATE(2024,1,15), 0) = DATE(2024,1,31)
        var startDate = new DateTime(2024, 1, 15).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(0),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        var expectedDate = new DateTime(2024, 1, 31).ToOADate();
        Assert.Equal(expectedDate, result.NumericValue);
    }

    [Fact]
    public void Eomonth_PositiveMonths_ReturnsEndOfFutureMonth()
    {
        var func = EomonthFunction.Instance;

        // EOMONTH(DATE(2024,1,15), 1) = DATE(2024,2,29)
        var startDate = new DateTime(2024, 1, 15).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(1),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        var expectedDate = new DateTime(2024, 2, 29).ToOADate(); // 2024 is leap year
        Assert.Equal(expectedDate, result.NumericValue);
    }

    [Fact]
    public void Eomonth_NegativeMonths_ReturnsEndOfPastMonth()
    {
        var func = EomonthFunction.Instance;

        // EOMONTH(DATE(2024,3,15), -1) = DATE(2024,2,29)
        var startDate = new DateTime(2024, 3, 15).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(-1),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        var expectedDate = new DateTime(2024, 2, 29).ToOADate();
        Assert.Equal(expectedDate, result.NumericValue);
    }

    [Fact]
    public void Eomonth_InvalidArguments_ReturnsError()
    {
        var func = EomonthFunction.Instance;

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
            CellValue.FromNumber(0),
        });

        Assert.True(result2.IsError);
        Assert.Equal("#VALUE!", result2.ErrorValue);
    }

    [Fact]
    public void Edate_ZeroMonths_ReturnsSameDate()
    {
        var func = EdateFunction.Instance;

        // EDATE(DATE(2024,1,15), 0) = DATE(2024,1,15)
        var startDate = new DateTime(2024, 1, 15).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(0),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(startDate, result.NumericValue);
    }

    [Fact]
    public void Edate_PositiveMonths_ReturnsFutureDate()
    {
        var func = EdateFunction.Instance;

        // EDATE(DATE(2024,1,15), 1) = DATE(2024,2,15)
        var startDate = new DateTime(2024, 1, 15).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(1),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        var expectedDate = new DateTime(2024, 2, 15).ToOADate();
        Assert.Equal(expectedDate, result.NumericValue);
    }

    [Fact]
    public void Edate_NegativeMonths_ReturnsPastDate()
    {
        var func = EdateFunction.Instance;

        // EDATE(DATE(2024,3,15), -1) = DATE(2024,2,15)
        var startDate = new DateTime(2024, 3, 15).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(-1),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        var expectedDate = new DateTime(2024, 2, 15).ToOADate();
        Assert.Equal(expectedDate, result.NumericValue);
    }

    [Fact]
    public void Edate_MultipleYears_ReturnsCorrectDate()
    {
        var func = EdateFunction.Instance;

        // EDATE(DATE(2024,1,15), 12) = DATE(2025,1,15)
        var startDate = new DateTime(2024, 1, 15).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(12),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        var expectedDate = new DateTime(2025, 1, 15).ToOADate();
        Assert.Equal(expectedDate, result.NumericValue);
    }

    [Fact]
    public void Edate_InvalidArguments_ReturnsError()
    {
        var func = EdateFunction.Instance;

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
            CellValue.FromNumber(1),
        });

        Assert.True(result2.IsError);
        Assert.Equal("#VALUE!", result2.ErrorValue);
    }

    [Fact]
    public void Days_ErrorPropagation_ReturnsError()
    {
        var func = DaysFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(44927),
        });

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void Time_ErrorPropagation_ReturnsError()
    {
        var func = TimeFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(12),
            CellValue.Error("#REF!"),
            CellValue.FromNumber(0),
        });

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    [Fact]
    public void Days360_ErrorPropagation_ReturnsError()
    {
        var func = Days360Function.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(44927),
            CellValue.FromNumber(44957),
            CellValue.Error("#N/A"),
        });

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }
}
