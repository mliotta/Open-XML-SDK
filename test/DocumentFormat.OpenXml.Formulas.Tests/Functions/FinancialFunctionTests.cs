// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for financial functions (PMT, FV, PV, NPER, RATE).
/// </summary>
public class FinancialFunctionTests
{
    // PMT Function Tests
    [Fact]
    public void Pmt_MonthlyLoanPayment_ReturnsCorrectValue()
    {
        // PMT(0.05/12, 360, 200000) - monthly payment on a 30-year mortgage
        var func = PmtFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.05 / 12),  // rate
            CellValue.FromNumber(360),         // nper
            CellValue.FromNumber(200000),      // pv
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(-1073.64, result.NumericValue, 2); // Expected: -1073.64
    }

    [Fact]
    public void Pmt_WithFutureValue_ReturnsCorrectValue()
    {
        // PMT(0.08/12, 10, 10000, 0, 1) - payment with future value
        var func = PmtFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.08 / 12),  // rate
            CellValue.FromNumber(10),          // nper
            CellValue.FromNumber(10000),       // pv
            CellValue.FromNumber(0),           // fv
            CellValue.FromNumber(1),           // type (beginning of period)
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(-1030.16, result.NumericValue, 2);
    }

    [Fact]
    public void Pmt_ZeroRate_ReturnsCorrectValue()
    {
        // PMT(0, 12, 12000) - zero interest payment
        var func = PmtFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0),       // rate
            CellValue.FromNumber(12),      // nper
            CellValue.FromNumber(12000),   // pv
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(-1000.0, result.NumericValue, 2);
    }

    [Fact]
    public void Pmt_InvalidArgCount_ReturnsError()
    {
        var func = PmtFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.05),
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Pmt_NegativeNper_ReturnsError()
    {
        var func = PmtFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.05),
            CellValue.FromNumber(-10),  // negative periods
            CellValue.FromNumber(10000),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Pmt_PropagatesError()
    {
        var func = PmtFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(10),
            CellValue.FromNumber(10000),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    // FV Function Tests
    [Fact]
    public void Fv_FutureValueOfInvestment_ReturnsCorrectValue()
    {
        // FV(0.06/12, 10, -200, -500, 1) - future value with monthly deposits
        var func = FvFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.06 / 12),  // rate
            CellValue.FromNumber(10),          // nper
            CellValue.FromNumber(-200),        // pmt (outflow)
            CellValue.FromNumber(-500),        // pv (outflow)
            CellValue.FromNumber(1),           // type
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(2581.40, result.NumericValue, 2);
    }

    [Fact]
    public void Fv_ZeroRate_ReturnsCorrectValue()
    {
        // FV(0, 10, -100, -1000) - zero interest
        var func = FvFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0),       // rate
            CellValue.FromNumber(10),      // nper
            CellValue.FromNumber(-100),    // pmt
            CellValue.FromNumber(-1000),   // pv
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(2000.0, result.NumericValue, 2);
    }

    [Fact]
    public void Fv_WithoutOptionalParams_ReturnsCorrectValue()
    {
        // FV(0.05/12, 60, -100) - future value with required params only
        var func = FvFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.05 / 12),  // rate
            CellValue.FromNumber(60),          // nper
            CellValue.FromNumber(-100),        // pmt
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue > 6000); // Approximate check
    }

    [Fact]
    public void Fv_InvalidType_ReturnsError()
    {
        var func = FvFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.05),
            CellValue.FromNumber(10),
            CellValue.FromNumber(-100),
            CellValue.FromNumber(0),
            CellValue.FromNumber(2),  // invalid type (must be 0 or 1)
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    // PV Function Tests
    [Fact]
    public void Pv_PresentValueOfLoan_ReturnsCorrectValue()
    {
        // PV(0.08/12, 240, 500) - present value of loan with monthly payments
        var func = PvFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.08 / 12),  // rate
            CellValue.FromNumber(240),         // nper
            CellValue.FromNumber(500),         // pmt
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(-59777.15, result.NumericValue, 2);
    }

    [Fact]
    public void Pv_WithFutureValue_ReturnsCorrectValue()
    {
        // PV(0.05/12, 60, -100, 10000, 0)
        var func = PvFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.05 / 12),  // rate
            CellValue.FromNumber(60),          // nper
            CellValue.FromNumber(-100),        // pmt
            CellValue.FromNumber(10000),       // fv
            CellValue.FromNumber(0),           // type
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue < 0); // Should be negative
    }

    [Fact]
    public void Pv_ZeroRate_ReturnsCorrectValue()
    {
        // PV(0, 12, -100, 5000) - zero interest
        var func = PvFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0),      // rate
            CellValue.FromNumber(12),     // nper
            CellValue.FromNumber(-100),   // pmt
            CellValue.FromNumber(5000),   // fv
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(-6200.0, result.NumericValue, 2);
    }

    [Fact]
    public void Pv_PropagatesError()
    {
        var func = PvFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.05),
            CellValue.Error("#N/A"),
            CellValue.FromNumber(-100),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    // NPER Function Tests
    [Fact]
    public void Nper_NumberOfPeriods_ReturnsCorrectValue()
    {
        // NPER(0.12/12, -100, -1000, 10000) - periods needed to grow investment
        var func = NperFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.12 / 12),  // rate
            CellValue.FromNumber(-100),        // pmt
            CellValue.FromNumber(-1000),       // pv
            CellValue.FromNumber(10000),       // fv
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(59.67, result.NumericValue, 2);
    }

    [Fact]
    public void Nper_ZeroRate_ReturnsCorrectValue()
    {
        // NPER(0, -100, -1000, 5000) - zero interest
        var func = NperFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0),      // rate
            CellValue.FromNumber(-100),   // pmt
            CellValue.FromNumber(-1000),  // pv
            CellValue.FromNumber(5000),   // fv
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(60.0, result.NumericValue, 2);
    }

    [Fact]
    public void Nper_WithoutOptionalParams_ReturnsCorrectValue()
    {
        // NPER(0.06/12, -100, 5000) - periods to pay off loan
        var func = NperFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.06 / 12),  // rate
            CellValue.FromNumber(-100),        // pmt
            CellValue.FromNumber(5000),        // pv
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue > 0);
    }

    [Fact]
    public void Nper_InvalidInputs_ReturnsError()
    {
        // Impossible scenario - trying to pay off loan with payments going the wrong way
        var func = NperFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.12 / 12),  // rate
            CellValue.FromNumber(100),         // pmt (wrong sign)
            CellValue.FromNumber(-1000),       // pv
            CellValue.FromNumber(10000),       // fv
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Nper_NegativeResult_ReturnsError()
    {
        var func = NperFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.12),
            CellValue.FromNumber(100),
            CellValue.FromNumber(10000),
            CellValue.FromNumber(5000),  // FV less than PV with positive payment
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    // RATE Function Tests
    [Fact]
    public void Rate_InterestRate_ReturnsCorrectValue()
    {
        // RATE(48, -200, 8000) - monthly rate for a loan
        var func = RateFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(48),      // nper
            CellValue.FromNumber(-200),    // pmt
            CellValue.FromNumber(8000),    // pv
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue > 0);
        Assert.True(result.NumericValue < 0.02); // Reasonable monthly rate
    }

    [Fact]
    public void Rate_WithAllParameters_ReturnsCorrectValue()
    {
        // RATE(12, -1000, 0, 13000, 1, 0.1)
        var func = RateFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(12),      // nper
            CellValue.FromNumber(-1000),   // pmt
            CellValue.FromNumber(0),       // pv
            CellValue.FromNumber(13000),   // fv
            CellValue.FromNumber(1),       // type
            CellValue.FromNumber(0.1),     // guess
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue > 0);
    }

    [Fact]
    public void Rate_ZeroRate_ReturnsZero()
    {
        // RATE(12, -1000, 0, 12000) - should return ~0 rate
        var func = RateFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(12),      // nper
            CellValue.FromNumber(-1000),   // pmt
            CellValue.FromNumber(0),       // pv
            CellValue.FromNumber(12000),   // fv
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(System.Math.Abs(result.NumericValue) < 0.001); // Close to zero
    }

    [Fact]
    public void Rate_InvalidNper_ReturnsError()
    {
        var func = RateFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-12),     // negative nper
            CellValue.FromNumber(-1000),
            CellValue.FromNumber(8000),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Rate_InvalidType_ReturnsError()
    {
        var func = RateFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(48),
            CellValue.FromNumber(-200),
            CellValue.FromNumber(8000),
            CellValue.FromNumber(0),
            CellValue.FromNumber(5),  // invalid type
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Rate_PropagatesError()
    {
        var func = RateFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(48),
            CellValue.Error("#REF!"),
            CellValue.FromNumber(8000),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    // Cross-function validation tests
    [Fact]
    public void Financial_PmtAndPvRelationship_IsConsistent()
    {
        // If we calculate PMT and then use it to calculate PV, we should get back original PV
        var rate = 0.06 / 12;
        var nper = 360;
        var pv = 100000;

        var pmtFunc = PmtFunction.Instance;
        var pmtArgs = new[]
        {
            CellValue.FromNumber(rate),
            CellValue.FromNumber(nper),
            CellValue.FromNumber(pv),
        };

        var pmtResult = pmtFunc.Execute(null!, pmtArgs);
        Assert.Equal(CellValueType.Number, pmtResult.Type);

        var pvFunc = PvFunction.Instance;
        var pvArgs = new[]
        {
            CellValue.FromNumber(rate),
            CellValue.FromNumber(nper),
            CellValue.FromNumber(pmtResult.NumericValue),
        };

        var pvResult = pvFunc.Execute(null!, pvArgs);
        Assert.Equal(CellValueType.Number, pvResult.Type);
        Assert.Equal(-pv, pvResult.NumericValue, 1); // Signs are opposite
    }

    [Fact]
    public void Financial_FvAndPvRelationship_IsConsistent()
    {
        // FV and PV should be inversely related
        var rate = 0.05 / 12;
        var nper = 60;
        var pmt = -100;

        var fvFunc = FvFunction.Instance;
        var fvArgs = new[]
        {
            CellValue.FromNumber(rate),
            CellValue.FromNumber(nper),
            CellValue.FromNumber(pmt),
        };

        var fvResult = fvFunc.Execute(null!, fvArgs);
        Assert.Equal(CellValueType.Number, fvResult.Type);

        // Now calculate PV using the FV we just calculated
        var pvFunc = PvFunction.Instance;
        var pvArgs = new[]
        {
            CellValue.FromNumber(rate),
            CellValue.FromNumber(nper),
            CellValue.FromNumber(pmt),
            CellValue.FromNumber(fvResult.NumericValue),
        };

        var pvResult = pvFunc.Execute(null!, pvArgs);
        Assert.Equal(CellValueType.Number, pvResult.Type);
        Assert.True(System.Math.Abs(pvResult.NumericValue) < 0.01); // Should be close to zero
    }
}
