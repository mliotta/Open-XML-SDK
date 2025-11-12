// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for YEARFRAC and DATEDIF date fraction calculation functions.
/// </summary>
public class DateFractionFunctionTests
{
    #region YEARFRAC Tests

    [Fact]
    public void Yearfrac_FullYear_Basis0_ReturnsOne()
    {
        var func = YearfracFunction.Instance;

        // YEARFRAC(DATE(2024,1,1), DATE(2024,12,31), 0) = 1.0 (30/360)
        var startDate = new DateTime(2024, 1, 1).ToOADate();
        var endDate = new DateTime(2024, 12, 31).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(endDate),
            CellValue.FromNumber(0),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue, 5);
    }

    [Fact]
    public void Yearfrac_HalfYear_Basis1_ReturnsApproximateHalf()
    {
        var func = YearfracFunction.Instance;

        // YEARFRAC(DATE(2024,1,1), DATE(2024,6,30), 1) ≈ 0.497 (actual/actual, 181 days in leap year)
        var startDate = new DateTime(2024, 1, 1).ToOADate();
        var endDate = new DateTime(2024, 6, 30).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(endDate),
            CellValue.FromNumber(1),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.InRange(result.NumericValue, 0.49, 0.50); // Approximately half a year
    }

    [Fact]
    public void Yearfrac_30Days_Basis2_ReturnsCorrectFraction()
    {
        var func = YearfracFunction.Instance;

        // YEARFRAC(DATE(2024,1,1), DATE(2024,1,31), 2) = 30/360 = 0.0833...
        var startDate = new DateTime(2024, 1, 1).ToOADate();
        var endDate = new DateTime(2024, 1, 31).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(endDate),
            CellValue.FromNumber(2),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(30.0 / 360.0, result.NumericValue, 5);
    }

    [Fact]
    public void Yearfrac_30Days_Basis3_ReturnsCorrectFraction()
    {
        var func = YearfracFunction.Instance;

        // YEARFRAC(DATE(2024,1,1), DATE(2024,1,31), 3) = 30/365 = 0.0821...
        var startDate = new DateTime(2024, 1, 1).ToOADate();
        var endDate = new DateTime(2024, 1, 31).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(endDate),
            CellValue.FromNumber(3),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(30.0 / 365.0, result.NumericValue, 5);
    }

    [Fact]
    public void Yearfrac_FullYear_Basis4_ReturnsOne()
    {
        var func = YearfracFunction.Instance;

        // YEARFRAC(DATE(2024,1,1), DATE(2024,12,31), 4) = 1.0 (European 30/360)
        var startDate = new DateTime(2024, 1, 1).ToOADate();
        var endDate = new DateTime(2024, 12, 31).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(endDate),
            CellValue.FromNumber(4),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue, 5);
    }

    [Fact]
    public void Yearfrac_DefaultBasis_UsesBasis0()
    {
        var func = YearfracFunction.Instance;

        // YEARFRAC(DATE(2024,1,1), DATE(2024,12,31)) defaults to basis 0
        var startDate = new DateTime(2024, 1, 1).ToOADate();
        var endDate = new DateTime(2024, 12, 31).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(endDate),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.0, result.NumericValue, 5);
    }

    [Fact]
    public void Yearfrac_StartGreaterThanEnd_ReturnsError()
    {
        var func = YearfracFunction.Instance;

        var startDate = new DateTime(2024, 12, 31).ToOADate();
        var endDate = new DateTime(2024, 1, 1).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(endDate),
            CellValue.FromNumber(0),
        });

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Yearfrac_InvalidBasis_ReturnsError()
    {
        var func = YearfracFunction.Instance;

        var startDate = new DateTime(2024, 1, 1).ToOADate();
        var endDate = new DateTime(2024, 12, 31).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(endDate),
            CellValue.FromNumber(5), // Invalid basis (must be 0-4)
        });

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Yearfrac_InvalidArguments_ReturnsError()
    {
        var func = YearfracFunction.Instance;

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
    public void Yearfrac_ErrorPropagation_ReturnsError()
    {
        var func = YearfracFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(44927),
            CellValue.FromNumber(0),
        });

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void Yearfrac_MultipleYears_Basis1_ReturnsCorrectFraction()
    {
        var func = YearfracFunction.Instance;

        // YEARFRAC(DATE(2023,1,1), DATE(2025,1,1), 1) = 2.0 (exactly 2 years)
        var startDate = new DateTime(2023, 1, 1).ToOADate();
        var endDate = new DateTime(2025, 1, 1).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(endDate),
            CellValue.FromNumber(1),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.InRange(result.NumericValue, 1.99, 2.01); // Very close to 2.0
    }

    #endregion

    #region DATEDIF Tests

    [Fact]
    public void Datedif_OneYear_UnitY_ReturnsOne()
    {
        var func = DatedifFunction.Instance;

        // DATEDIF(DATE(2023,1,1), DATE(2024,1,1), "Y") = 1
        var startDate = new DateTime(2023, 1, 1).ToOADate();
        var endDate = new DateTime(2024, 1, 1).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(endDate),
            CellValue.FromString("Y"),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1, result.NumericValue);
    }

    [Fact]
    public void Datedif_TwoMonths_UnitM_ReturnsTwo()
    {
        var func = DatedifFunction.Instance;

        // DATEDIF(DATE(2023,1,1), DATE(2023,3,1), "M") = 2
        var startDate = new DateTime(2023, 1, 1).ToOADate();
        var endDate = new DateTime(2023, 3, 1).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(endDate),
            CellValue.FromString("M"),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(2, result.NumericValue);
    }

    [Fact]
    public void Datedif_FourteenDays_UnitD_ReturnsFourteen()
    {
        var func = DatedifFunction.Instance;

        // DATEDIF(DATE(2023,1,1), DATE(2023,1,15), "D") = 14
        var startDate = new DateTime(2023, 1, 1).ToOADate();
        var endDate = new DateTime(2023, 1, 15).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(endDate),
            CellValue.FromString("D"),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(14, result.NumericValue);
    }

    [Fact]
    public void Datedif_MonthsExcludingYears_UnitYM_ReturnsCorrect()
    {
        var func = DatedifFunction.Instance;

        // DATEDIF(DATE(2023,1,1), DATE(2024,3,1), "YM") = 2 (months excluding years)
        var startDate = new DateTime(2023, 1, 1).ToOADate();
        var endDate = new DateTime(2024, 3, 1).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(endDate),
            CellValue.FromString("YM"),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(2, result.NumericValue);
    }

    [Fact]
    public void Datedif_DaysExcludingYears_UnitYD_ReturnsCorrect()
    {
        var func = DatedifFunction.Instance;

        // DATEDIF(DATE(2023,1,1), DATE(2024,1,15), "YD") = 14 (days excluding years)
        var startDate = new DateTime(2023, 1, 1).ToOADate();
        var endDate = new DateTime(2024, 1, 15).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(endDate),
            CellValue.FromString("YD"),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(14, result.NumericValue);
    }

    [Fact]
    public void Datedif_DaysExcludingMonthsAndYears_UnitMD_ReturnsCorrect()
    {
        var func = DatedifFunction.Instance;

        // DATEDIF(DATE(2023,1,1), DATE(2023,2,15), "MD") = 14 (days excluding months)
        var startDate = new DateTime(2023, 1, 1).ToOADate();
        var endDate = new DateTime(2023, 2, 15).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(endDate),
            CellValue.FromString("MD"),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(14, result.NumericValue);
    }

    [Fact]
    public void Datedif_CaseInsensitive_WorksCorrectly()
    {
        var func = DatedifFunction.Instance;

        var startDate = new DateTime(2023, 1, 1).ToOADate();
        var endDate = new DateTime(2024, 1, 1).ToOADate();

        // Test lowercase
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(endDate),
            CellValue.FromString("y"),
        });

        Assert.Equal(CellValueType.Number, result1.Type);
        Assert.Equal(1, result1.NumericValue);

        // Test mixed case
        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(endDate),
            CellValue.FromString("Y"),
        });

        Assert.Equal(CellValueType.Number, result2.Type);
        Assert.Equal(1, result2.NumericValue);
    }

    [Fact]
    public void Datedif_StartGreaterThanEnd_ReturnsError()
    {
        var func = DatedifFunction.Instance;

        var startDate = new DateTime(2024, 1, 1).ToOADate();
        var endDate = new DateTime(2023, 1, 1).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(endDate),
            CellValue.FromString("Y"),
        });

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Datedif_InvalidUnit_ReturnsError()
    {
        var func = DatedifFunction.Instance;

        var startDate = new DateTime(2023, 1, 1).ToOADate();
        var endDate = new DateTime(2024, 1, 1).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(endDate),
            CellValue.FromString("X"), // Invalid unit
        });

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Datedif_InvalidArguments_ReturnsError()
    {
        var func = DatedifFunction.Instance;

        // Wrong number of arguments
        var result1 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(44927),
            CellValue.FromNumber(44957),
        });

        Assert.True(result1.IsError);
        Assert.Equal("#VALUE!", result1.ErrorValue);

        // Non-text unit argument
        var result2 = func.Execute(null!, new[]
        {
            CellValue.FromNumber(44927),
            CellValue.FromNumber(44957),
            CellValue.FromNumber(1), // Should be text
        });

        Assert.True(result2.IsError);
        Assert.Equal("#VALUE!", result2.ErrorValue);
    }

    [Fact]
    public void Datedif_ErrorPropagation_ReturnsError()
    {
        var func = DatedifFunction.Instance;

        var result = func.Execute(null!, new[]
        {
            CellValue.Error("#REF!"),
            CellValue.FromNumber(44927),
            CellValue.FromString("Y"),
        });

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    [Fact]
    public void Datedif_LeapYear_HandlesCorrectly()
    {
        var func = DatedifFunction.Instance;

        // DATEDIF(DATE(2024,2,28), DATE(2024,3,1), "D") = 2 (leap year)
        var startDate = new DateTime(2024, 2, 28).ToOADate();
        var endDate = new DateTime(2024, 3, 1).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(endDate),
            CellValue.FromString("D"),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(2, result.NumericValue); // Feb 28 -> Feb 29 -> Mar 1 = 2 days
    }

    [Fact]
    public void Datedif_IncompleteYear_UnitY_ReturnsZero()
    {
        var func = DatedifFunction.Instance;

        // DATEDIF(DATE(2023,1,1), DATE(2023,12,31), "Y") = 0 (not a full year)
        var startDate = new DateTime(2023, 1, 1).ToOADate();
        var endDate = new DateTime(2023, 12, 31).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(endDate),
            CellValue.FromString("Y"),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0, result.NumericValue);
    }

    [Fact]
    public void Datedif_ComplexYM_CrossingYears_ReturnsCorrect()
    {
        var func = DatedifFunction.Instance;

        // DATEDIF(DATE(2023,10,1), DATE(2024,3,15), "YM") = 5 (Oct -> Mar = 5 months)
        var startDate = new DateTime(2023, 10, 1).ToOADate();
        var endDate = new DateTime(2024, 3, 15).ToOADate();

        var result = func.Execute(null!, new[]
        {
            CellValue.FromNumber(startDate),
            CellValue.FromNumber(endDate),
            CellValue.FromString("YM"),
        });

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(5, result.NumericValue);
    }

    #endregion
}
