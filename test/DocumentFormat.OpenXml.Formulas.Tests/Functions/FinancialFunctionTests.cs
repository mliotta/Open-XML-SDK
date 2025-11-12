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

    // NPV Function Tests
    [Fact]
    public void Npv_BasicCalculation_ReturnsCorrectValue()
    {
        // NPV(0.10, -10000, 3000, 4200, 6800)
        var func = NpvFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.10),      // rate
            CellValue.FromNumber(-10000),    // value1
            CellValue.FromNumber(3000),      // value2
            CellValue.FromNumber(4200),      // value3
            CellValue.FromNumber(6800),      // value4
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1188.44, result.NumericValue, 2);
    }

    [Fact]
    public void Npv_SingleValue_ReturnsCorrectValue()
    {
        // NPV(0.05, 1000)
        var func = NpvFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.05),   // rate
            CellValue.FromNumber(1000),   // value1
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(952.38, result.NumericValue, 2); // 1000 / 1.05
    }

    [Fact]
    public void Npv_ZeroRate_ReturnsSum()
    {
        // NPV(0, 100, 200, 300) - with zero rate, NPV is simple sum
        var func = NpvFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0),      // rate
            CellValue.FromNumber(100),
            CellValue.FromNumber(200),
            CellValue.FromNumber(300),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(600.0, result.NumericValue, 2);
    }

    [Fact]
    public void Npv_InvalidArgCount_ReturnsError()
    {
        var func = NpvFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.10),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Npv_PropagatesError()
    {
        var func = NpvFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.10),
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(100),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    // IRR Function Tests
    [Fact]
    public void Irr_BasicCalculation_ReturnsCorrectValue()
    {
        // IRR(-10000, 3000, 4200, 6800) - should match NPV example
        var func = IrrFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-10000),
            CellValue.FromNumber(3000),
            CellValue.FromNumber(4200),
            CellValue.FromNumber(6800),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue > 0.1); // Should be positive rate
        Assert.True(result.NumericValue < 0.3); // Reasonable range
    }

    [Fact]
    public void Irr_SimpleInvestment_ReturnsCorrectValue()
    {
        // IRR(-1000, 300, 400, 500) - simple investment
        var func = IrrFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-1000),
            CellValue.FromNumber(300),
            CellValue.FromNumber(400),
            CellValue.FromNumber(500),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue > 0); // Positive return
    }

    [Fact]
    public void Irr_AllPositive_ReturnsError()
    {
        // IRR requires both positive and negative values
        var func = IrrFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(100),
            CellValue.FromNumber(200),
            CellValue.FromNumber(300),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Irr_AllNegative_ReturnsError()
    {
        var func = IrrFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-100),
            CellValue.FromNumber(-200),
            CellValue.FromNumber(-300),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Irr_PropagatesError()
    {
        var func = IrrFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-1000),
            CellValue.Error("#N/A"),
            CellValue.FromNumber(500),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    // IPMT Function Tests
    [Fact]
    public void Ipmt_FirstPaymentInterest_ReturnsCorrectValue()
    {
        // IPMT(0.10/12, 1, 360, 200000) - interest portion of first payment on 30-year loan
        var func = IpmtFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.10 / 12),  // rate
            CellValue.FromNumber(1),          // per
            CellValue.FromNumber(360),        // nper
            CellValue.FromNumber(200000),     // pv
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue < 0); // Interest is negative (outflow)
        Assert.True(System.Math.Abs(result.NumericValue) > 1600); // Approximate check
    }

    [Fact]
    public void Ipmt_MiddlePaymentInterest_ReturnsCorrectValue()
    {
        // IPMT(0.10/12, 180, 360, 200000) - interest at midpoint
        var func = IpmtFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.10 / 12),  // rate
            CellValue.FromNumber(180),        // per
            CellValue.FromNumber(360),        // nper
            CellValue.FromNumber(200000),     // pv
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue < 0); // Interest is negative
    }

    [Fact]
    public void Ipmt_WithFutureValue_ReturnsCorrectValue()
    {
        // IPMT(0.08/12, 5, 60, 10000, 5000, 0)
        var func = IpmtFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.08 / 12),  // rate
            CellValue.FromNumber(5),          // per
            CellValue.FromNumber(60),         // nper
            CellValue.FromNumber(10000),      // pv
            CellValue.FromNumber(5000),       // fv
            CellValue.FromNumber(0),          // type
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue < 0); // Interest is negative
    }

    [Fact]
    public void Ipmt_BeginningOfPeriod_ReturnsCorrectValue()
    {
        // IPMT(0.08/12, 1, 60, 10000, 0, 1) - beginning of period
        var func = IpmtFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.08 / 12),  // rate
            CellValue.FromNumber(1),          // per
            CellValue.FromNumber(60),         // nper
            CellValue.FromNumber(10000),      // pv
            CellValue.FromNumber(0),          // fv
            CellValue.FromNumber(1),          // type
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // For beginning of period, first payment has no interest
        Assert.Equal(0.0, result.NumericValue, 2);
    }

    [Fact]
    public void Ipmt_InvalidPeriod_ReturnsError()
    {
        var func = IpmtFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.10 / 12),
            CellValue.FromNumber(0),          // invalid period (< 1)
            CellValue.FromNumber(360),
            CellValue.FromNumber(200000),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Ipmt_PeriodExceedsNper_ReturnsError()
    {
        var func = IpmtFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.10 / 12),
            CellValue.FromNumber(361),        // period > nper
            CellValue.FromNumber(360),
            CellValue.FromNumber(200000),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Ipmt_PropagatesError()
    {
        var func = IpmtFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#REF!"),
            CellValue.FromNumber(1),
            CellValue.FromNumber(360),
            CellValue.FromNumber(200000),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    // PPMT Function Tests
    [Fact]
    public void Ppmt_FirstPaymentPrincipal_ReturnsCorrectValue()
    {
        // PPMT(0.10/12, 1, 360, 200000) - principal portion of first payment
        var func = PpmtFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.10 / 12),  // rate
            CellValue.FromNumber(1),          // per
            CellValue.FromNumber(360),        // nper
            CellValue.FromNumber(200000),     // pv
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue < 0); // Principal is negative (outflow)
    }

    [Fact]
    public void Ppmt_LastPaymentPrincipal_ReturnsCorrectValue()
    {
        // PPMT(0.10/12, 360, 360, 200000) - last payment has mostly principal
        var func = PpmtFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.10 / 12),  // rate
            CellValue.FromNumber(360),        // per
            CellValue.FromNumber(360),        // nper
            CellValue.FromNumber(200000),     // pv
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue < 0); // Principal is negative
    }

    [Fact]
    public void Ppmt_WithFutureValue_ReturnsCorrectValue()
    {
        // PPMT(0.08/12, 5, 60, 10000, 5000, 0)
        var func = PpmtFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.08 / 12),  // rate
            CellValue.FromNumber(5),          // per
            CellValue.FromNumber(60),         // nper
            CellValue.FromNumber(10000),      // pv
            CellValue.FromNumber(5000),       // fv
            CellValue.FromNumber(0),          // type
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue < 0); // Principal is negative
    }

    [Fact]
    public void Ppmt_BeginningOfPeriod_ReturnsCorrectValue()
    {
        // PPMT(0.08/12, 1, 60, 10000, 0, 1) - beginning of period
        var func = PpmtFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.08 / 12),  // rate
            CellValue.FromNumber(1),          // per
            CellValue.FromNumber(60),         // nper
            CellValue.FromNumber(10000),      // pv
            CellValue.FromNumber(0),          // fv
            CellValue.FromNumber(1),          // type
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue < 0); // Principal is negative
    }

    [Fact]
    public void Ppmt_InvalidPeriod_ReturnsError()
    {
        var func = PpmtFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.10 / 12),
            CellValue.FromNumber(0),          // invalid period (< 1)
            CellValue.FromNumber(360),
            CellValue.FromNumber(200000),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Ppmt_PeriodExceedsNper_ReturnsError()
    {
        var func = PpmtFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.10 / 12),
            CellValue.FromNumber(361),        // period > nper
            CellValue.FromNumber(360),
            CellValue.FromNumber(200000),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Ppmt_PropagatesError()
    {
        var func = PpmtFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.10 / 12),
            CellValue.FromNumber(1),
            CellValue.Error("#N/A"),
            CellValue.FromNumber(200000),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
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

    [Fact]
    public void Financial_IpmtAndPpmtSumToPmt_IsConsistent()
    {
        // IPMT + PPMT should equal PMT for any period
        var rate = 0.10 / 12;
        var per = 10;
        var nper = 360;
        var pv = 200000;

        var pmtFunc = PmtFunction.Instance;
        var pmtArgs = new[]
        {
            CellValue.FromNumber(rate),
            CellValue.FromNumber(nper),
            CellValue.FromNumber(pv),
        };

        var pmtResult = pmtFunc.Execute(null!, pmtArgs);
        Assert.Equal(CellValueType.Number, pmtResult.Type);

        var ipmtFunc = IpmtFunction.Instance;
        var ipmtArgs = new[]
        {
            CellValue.FromNumber(rate),
            CellValue.FromNumber(per),
            CellValue.FromNumber(nper),
            CellValue.FromNumber(pv),
        };

        var ipmtResult = ipmtFunc.Execute(null!, ipmtArgs);
        Assert.Equal(CellValueType.Number, ipmtResult.Type);

        var ppmtFunc = PpmtFunction.Instance;
        var ppmtArgs = new[]
        {
            CellValue.FromNumber(rate),
            CellValue.FromNumber(per),
            CellValue.FromNumber(nper),
            CellValue.FromNumber(pv),
        };

        var ppmtResult = ppmtFunc.Execute(null!, ppmtArgs);
        Assert.Equal(CellValueType.Number, ppmtResult.Type);

        // IPMT + PPMT should equal PMT
        var sum = ipmtResult.NumericValue + ppmtResult.NumericValue;
        Assert.Equal(pmtResult.NumericValue, sum, 2);
    }

    [Fact]
    public void Financial_IrrAndNpvRelationship_IsConsistent()
    {
        // NPV at IRR should be close to zero
        var values = new[]
        {
            -10000.0,
            3000.0,
            4200.0,
            6800.0,
        };

        var irrFunc = IrrFunction.Instance;
        var irrArgs = new[]
        {
            CellValue.FromNumber(values[0]),
            CellValue.FromNumber(values[1]),
            CellValue.FromNumber(values[2]),
            CellValue.FromNumber(values[3]),
        };

        var irrResult = irrFunc.Execute(null!, irrArgs);
        Assert.Equal(CellValueType.Number, irrResult.Type);

        // Now calculate NPV at the IRR rate
        var npvFunc = NpvFunction.Instance;
        var npvArgs = new[]
        {
            CellValue.FromNumber(irrResult.NumericValue),
            CellValue.FromNumber(values[1]),
            CellValue.FromNumber(values[2]),
            CellValue.FromNumber(values[3]),
        };

        var npvResult = npvFunc.Execute(null!, npvArgs);
        Assert.Equal(CellValueType.Number, npvResult.Type);

        // NPV at IRR plus the initial investment should be close to zero
        var totalNpv = npvResult.NumericValue + values[0];
        Assert.True(System.Math.Abs(totalNpv) < 1.0); // Close to zero
    }

    // SLN Function Tests
    [Fact]
    public void Sln_BasicCalculation_ReturnsCorrectValue()
    {
        // SLN(30000, 7500, 10) - straight-line depreciation
        var func = SlnFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(30000),  // cost
            CellValue.FromNumber(7500),   // salvage
            CellValue.FromNumber(10),     // life
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(2250.0, result.NumericValue, 2);
    }

    [Fact]
    public void Sln_ZeroSalvage_ReturnsCorrectValue()
    {
        // SLN(10000, 0, 5)
        var func = SlnFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10000),
            CellValue.FromNumber(0),
            CellValue.FromNumber(5),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(2000.0, result.NumericValue, 2);
    }

    [Fact]
    public void Sln_InvalidLife_ReturnsError()
    {
        var func = SlnFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(30000),
            CellValue.FromNumber(7500),
            CellValue.FromNumber(0),  // invalid life
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Sln_InvalidArgCount_ReturnsError()
    {
        var func = SlnFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(30000),
            CellValue.FromNumber(7500),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    [Fact]
    public void Sln_PropagatesError()
    {
        var func = SlnFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(7500),
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    // DB Function Tests
    [Fact]
    public void Db_FirstPeriod_ReturnsCorrectValue()
    {
        // DB(1000000, 100000, 6, 1) - first period declining balance
        var func = DbFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1000000),  // cost
            CellValue.FromNumber(100000),   // salvage
            CellValue.FromNumber(6),        // life
            CellValue.FromNumber(1),        // period
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue > 0);
        Assert.True(result.NumericValue < 1000000);
    }

    [Fact]
    public void Db_WithMonthParameter_ReturnsCorrectValue()
    {
        // DB(1000000, 100000, 6, 1, 7) - first period with 7 months
        var func = DbFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1000000),
            CellValue.FromNumber(100000),
            CellValue.FromNumber(6),
            CellValue.FromNumber(1),
            CellValue.FromNumber(7),  // 7 months in first year
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue > 0);
    }

    [Fact]
    public void Db_LastPeriod_ReturnsCorrectValue()
    {
        // DB(1000000, 100000, 6, 6) - last period
        var func = DbFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1000000),
            CellValue.FromNumber(100000),
            CellValue.FromNumber(6),
            CellValue.FromNumber(6),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue > 0);
    }

    [Fact]
    public void Db_InvalidPeriod_ReturnsError()
    {
        var func = DbFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1000000),
            CellValue.FromNumber(100000),
            CellValue.FromNumber(6),
            CellValue.FromNumber(0),  // invalid period
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Db_PeriodExceedsLife_ReturnsError()
    {
        var func = DbFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1000000),
            CellValue.FromNumber(100000),
            CellValue.FromNumber(6),
            CellValue.FromNumber(7),  // period > life
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Db_SalvageGreaterThanCost_ReturnsZero()
    {
        var func = DbFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(100000),
            CellValue.FromNumber(200000),  // salvage > cost
            CellValue.FromNumber(6),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.0, result.NumericValue, 2);
    }

    [Fact]
    public void Db_PropagatesError()
    {
        var func = DbFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1000000),
            CellValue.Error("#REF!"),
            CellValue.FromNumber(6),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }

    // DDB Function Tests
    [Fact]
    public void Ddb_FirstPeriod_ReturnsCorrectValue()
    {
        // DDB(2400, 300, 10, 1) - first period double-declining balance
        var func = DdbFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(2400),  // cost
            CellValue.FromNumber(300),   // salvage
            CellValue.FromNumber(10),    // life
            CellValue.FromNumber(1),     // period
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(480.0, result.NumericValue, 2); // 2400 * 0.2
    }

    [Fact]
    public void Ddb_SecondPeriod_ReturnsCorrectValue()
    {
        // DDB(2400, 300, 10, 2)
        var func = DdbFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(2400),
            CellValue.FromNumber(300),
            CellValue.FromNumber(10),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(384.0, result.NumericValue, 2); // (2400 - 480) * 0.2
    }

    [Fact]
    public void Ddb_WithCustomFactor_ReturnsCorrectValue()
    {
        // DDB(2400, 300, 10, 1, 1.5) - with factor of 1.5
        var func = DdbFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(2400),
            CellValue.FromNumber(300),
            CellValue.FromNumber(10),
            CellValue.FromNumber(1),
            CellValue.FromNumber(1.5),  // factor
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(360.0, result.NumericValue, 2); // 2400 * 0.15
    }

    [Fact]
    public void Ddb_LastPeriod_ReturnsCorrectValue()
    {
        // DDB(2400, 300, 10, 10)
        var func = DdbFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(2400),
            CellValue.FromNumber(300),
            CellValue.FromNumber(10),
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue >= 0);
    }

    [Fact]
    public void Ddb_InvalidPeriod_ReturnsError()
    {
        var func = DdbFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(2400),
            CellValue.FromNumber(300),
            CellValue.FromNumber(10),
            CellValue.FromNumber(0),  // invalid period
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Ddb_InvalidFactor_ReturnsError()
    {
        var func = DdbFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(2400),
            CellValue.FromNumber(300),
            CellValue.FromNumber(10),
            CellValue.FromNumber(1),
            CellValue.FromNumber(-1),  // invalid factor
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Ddb_SalvageGreaterThanCost_ReturnsZero()
    {
        var func = DdbFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1000),
            CellValue.FromNumber(2000),  // salvage > cost
            CellValue.FromNumber(10),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.0, result.NumericValue, 2);
    }

    [Fact]
    public void Ddb_PropagatesError()
    {
        var func = DdbFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(2400),
            CellValue.FromNumber(300),
            CellValue.Error("#N/A"),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    // SYD Function Tests
    [Fact]
    public void Syd_FirstPeriod_ReturnsCorrectValue()
    {
        // SYD(30000, 7500, 10, 1) - first period sum-of-years' digits
        var func = SydFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(30000),  // cost
            CellValue.FromNumber(7500),   // salvage
            CellValue.FromNumber(10),     // life
            CellValue.FromNumber(1),      // period
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // Formula: (30000 - 7500) * (10 - 1 + 1) * 2 / (10 * 11) = 22500 * 10 * 2 / 110 = 4090.91
        Assert.Equal(4090.91, result.NumericValue, 2);
    }

    [Fact]
    public void Syd_SecondPeriod_ReturnsCorrectValue()
    {
        // SYD(30000, 7500, 10, 2)
        var func = SydFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(30000),
            CellValue.FromNumber(7500),
            CellValue.FromNumber(10),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // Formula: (30000 - 7500) * 9 * 2 / 110 = 3681.82
        Assert.Equal(3681.82, result.NumericValue, 2);
    }

    [Fact]
    public void Syd_LastPeriod_ReturnsCorrectValue()
    {
        // SYD(30000, 7500, 10, 10)
        var func = SydFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(30000),
            CellValue.FromNumber(7500),
            CellValue.FromNumber(10),
            CellValue.FromNumber(10),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // Formula: (30000 - 7500) * 1 * 2 / 110 = 409.09
        Assert.Equal(409.09, result.NumericValue, 2);
    }

    [Fact]
    public void Syd_ZeroSalvage_ReturnsCorrectValue()
    {
        // SYD(10000, 0, 5, 1)
        var func = SydFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10000),
            CellValue.FromNumber(0),
            CellValue.FromNumber(5),
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // Formula: 10000 * 5 * 2 / 30 = 3333.33
        Assert.Equal(3333.33, result.NumericValue, 2);
    }

    [Fact]
    public void Syd_InvalidPeriod_ReturnsError()
    {
        var func = SydFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(30000),
            CellValue.FromNumber(7500),
            CellValue.FromNumber(10),
            CellValue.FromNumber(0),  // invalid period
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Syd_PeriodExceedsLife_ReturnsError()
    {
        var func = SydFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(30000),
            CellValue.FromNumber(7500),
            CellValue.FromNumber(10),
            CellValue.FromNumber(11),  // period > life
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Syd_InvalidLife_ReturnsError()
    {
        var func = SydFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(30000),
            CellValue.FromNumber(7500),
            CellValue.FromNumber(0),  // invalid life
            CellValue.FromNumber(1),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Syd_PropagatesError()
    {
        var func = SydFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(30000),
            CellValue.FromNumber(7500),
            CellValue.FromNumber(10),
            CellValue.Error("#VALUE!"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    // Cross-function validation for depreciation methods
    [Fact]
    public void Depreciation_TotalSlnDepreciation_EqualsDepreciableAmount()
    {
        // Sum of all SLN depreciation should equal depreciable amount
        var cost = 30000.0;
        var salvage = 7500.0;
        var life = 10.0;
        var func = SlnFunction.Instance;

        double totalDepreciation = 0.0;
        for (int period = 1; period <= life; period++)
        {
            var args = new[]
            {
                CellValue.FromNumber(cost),
                CellValue.FromNumber(salvage),
                CellValue.FromNumber(life),
            };

            var result = func.Execute(null!, args);
            totalDepreciation += result.NumericValue;
        }

        Assert.Equal(cost - salvage, totalDepreciation, 1);
    }

    [Fact]
    public void Depreciation_TotalSydDepreciation_EqualsDepreciableAmount()
    {
        // Sum of all SYD depreciation should equal depreciable amount
        var cost = 30000.0;
        var salvage = 7500.0;
        var life = 10.0;
        var func = SydFunction.Instance;

        double totalDepreciation = 0.0;
        for (int period = 1; period <= life; period++)
        {
            var args = new[]
            {
                CellValue.FromNumber(cost),
                CellValue.FromNumber(salvage),
                CellValue.FromNumber(life),
                CellValue.FromNumber(period),
            };

            var result = func.Execute(null!, args);
            totalDepreciation += result.NumericValue;
        }

        Assert.Equal(cost - salvage, totalDepreciation, 1);
    }

    [Fact]
    public void Depreciation_SydFirstPeriodGreaterThanSln_IsConsistent()
    {
        // SYD should have higher depreciation in first period than SLN
        var cost = 30000.0;
        var salvage = 7500.0;
        var life = 10.0;

        var slnFunc = SlnFunction.Instance;
        var slnArgs = new[]
        {
            CellValue.FromNumber(cost),
            CellValue.FromNumber(salvage),
            CellValue.FromNumber(life),
        };
        var slnResult = slnFunc.Execute(null!, slnArgs);

        var sydFunc = SydFunction.Instance;
        var sydArgs = new[]
        {
            CellValue.FromNumber(cost),
            CellValue.FromNumber(salvage),
            CellValue.FromNumber(life),
            CellValue.FromNumber(1),
        };
        var sydResult = sydFunc.Execute(null!, sydArgs);

        Assert.True(sydResult.NumericValue > slnResult.NumericValue);
    }

    [Fact]
    public void Depreciation_DdbFirstPeriodGreaterThanSln_IsConsistent()
    {
        // DDB should have higher depreciation in first period than SLN
        var cost = 2400.0;
        var salvage = 300.0;
        var life = 10.0;

        var slnFunc = SlnFunction.Instance;
        var slnArgs = new[]
        {
            CellValue.FromNumber(cost),
            CellValue.FromNumber(salvage),
            CellValue.FromNumber(life),
        };
        var slnResult = slnFunc.Execute(null!, slnArgs);

        var ddbFunc = DdbFunction.Instance;
        var ddbArgs = new[]
        {
            CellValue.FromNumber(cost),
            CellValue.FromNumber(salvage),
            CellValue.FromNumber(life),
            CellValue.FromNumber(1),
        };
        var ddbResult = ddbFunc.Execute(null!, ddbArgs);

        Assert.True(ddbResult.NumericValue > slnResult.NumericValue);
    }
}
