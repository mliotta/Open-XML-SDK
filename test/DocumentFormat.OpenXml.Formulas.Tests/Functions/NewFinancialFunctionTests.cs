// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for newly implemented financial functions.
/// </summary>
public class NewFinancialFunctionTests
{
    // EFFECT Function Tests
    [Fact]
    public void Effect_QuarterlyCompounding_ReturnsCorrectValue()
    {
        // EFFECT(0.05, 4) - 5% nominal rate compounded quarterly
        var func = EffectFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.05),  // nominal rate
            CellValue.FromNumber(4),      // periods per year
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.05095, result.NumericValue, 5);
    }

    [Fact]
    public void Effect_MonthlyCompounding_ReturnsCorrectValue()
    {
        // EFFECT(0.08, 12) - 8% nominal rate compounded monthly
        var func = EffectFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.08),
            CellValue.FromNumber(12),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.08300, result.NumericValue, 5);
    }

    [Fact]
    public void Effect_NegativeRate_ReturnsError()
    {
        var func = EffectFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-0.05),
            CellValue.FromNumber(4),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Effect_InvalidPeriods_ReturnsError()
    {
        var func = EffectFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.05),
            CellValue.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    // NOMINAL Function Tests
    [Fact]
    public void Nominal_QuarterlyCompounding_ReturnsCorrectValue()
    {
        // NOMINAL(0.05095, 4) - effective rate to nominal rate
        var func = NominalFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.05095),
            CellValue.FromNumber(4),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.05, result.NumericValue, 4);
    }

    [Fact]
    public void Nominal_MonthlyCompounding_ReturnsCorrectValue()
    {
        // NOMINAL(0.083, 12)
        var func = NominalFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.083),
            CellValue.FromNumber(12),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(0.08, result.NumericValue, 3);
    }

    [Fact]
    public void Nominal_EffectAndNominalInverse_ReturnsOriginalValue()
    {
        // Test that NOMINAL(EFFECT(rate, n), n) = rate
        var effectFunc = EffectFunction.Instance;
        var nominalFunc = NominalFunction.Instance;

        var rate = 0.06;
        var n = 12.0;

        var effectArgs = new[]
        {
            CellValue.FromNumber(rate),
            CellValue.FromNumber(n),
        };

        var effectResult = effectFunc.Execute(null!, effectArgs);
        Assert.Equal(CellValueType.Number, effectResult.Type);

        var nominalArgs = new[]
        {
            effectResult,
            CellValue.FromNumber(n),
        };

        var nominalResult = nominalFunc.Execute(null!, nominalArgs);
        Assert.Equal(CellValueType.Number, nominalResult.Type);
        Assert.Equal(rate, nominalResult.NumericValue, 6);
    }

    // MIRR Function Tests
    [Fact]
    public void Mirr_BasicCalculation_ReturnsCorrectValue()
    {
        // MIRR(-10000, 3000, 4200, 6800, 0.1, 0.12)
        var func = MirrFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-10000),
            CellValue.FromNumber(3000),
            CellValue.FromNumber(4200),
            CellValue.FromNumber(6800),
            CellValue.FromNumber(0.1),   // finance rate
            CellValue.FromNumber(0.12),  // reinvest rate
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue > 0.1);
        Assert.True(result.NumericValue < 0.3);
    }

    [Fact]
    public void Mirr_AllPositive_ReturnsError()
    {
        var func = MirrFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1000),
            CellValue.FromNumber(2000),
            CellValue.FromNumber(3000),
            CellValue.FromNumber(0.1),
            CellValue.FromNumber(0.12),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void Mirr_AllNegative_ReturnsError()
    {
        var func = MirrFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-1000),
            CellValue.FromNumber(-2000),
            CellValue.FromNumber(-3000),
            CellValue.FromNumber(0.1),
            CellValue.FromNumber(0.12),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    // CUMIPMT Function Tests
    [Fact]
    public void Cumipmt_FirstYearInterest_ReturnsCorrectValue()
    {
        // CUMIPMT(0.06/12, 360, 100000, 1, 12, 0) - first year interest on 30-year mortgage
        var func = CumipmtFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.06 / 12),  // rate
            CellValue.FromNumber(360),        // nper
            CellValue.FromNumber(100000),     // pv
            CellValue.FromNumber(1),          // start_period
            CellValue.FromNumber(12),         // end_period
            CellValue.FromNumber(0),          // type
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue < 0); // Interest is negative (outflow)
        Assert.True(Math.Abs(result.NumericValue) > 5000); // Approximate check
    }

    [Fact]
    public void Cumipmt_SinglePeriod_ReturnsCorrectValue()
    {
        // CUMIPMT for a single period should match IPMT
        var func = CumipmtFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.08 / 12),
            CellValue.FromNumber(360),
            CellValue.FromNumber(100000),
            CellValue.FromNumber(5),
            CellValue.FromNumber(5),  // Same start and end
            CellValue.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue < 0);
    }

    [Fact]
    public void Cumipmt_InvalidPeriodRange_ReturnsError()
    {
        var func = CumipmtFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.06 / 12),
            CellValue.FromNumber(360),
            CellValue.FromNumber(100000),
            CellValue.FromNumber(12),  // start > end
            CellValue.FromNumber(1),
            CellValue.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    // CUMPRINC Function Tests
    [Fact]
    public void Cumprinc_FirstYearPrincipal_ReturnsCorrectValue()
    {
        // CUMPRINC(0.06/12, 360, 100000, 1, 12, 0)
        var func = CumprincFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.06 / 12),
            CellValue.FromNumber(360),
            CellValue.FromNumber(100000),
            CellValue.FromNumber(1),
            CellValue.FromNumber(12),
            CellValue.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue < 0); // Principal is negative (outflow)
    }

    [Fact]
    public void Cumprinc_SinglePeriod_ReturnsCorrectValue()
    {
        var func = CumprincFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.08 / 12),
            CellValue.FromNumber(360),
            CellValue.FromNumber(100000),
            CellValue.FromNumber(5),
            CellValue.FromNumber(5),
            CellValue.FromNumber(0),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue < 0);
    }

    [Fact]
    public void CumprincAndCumipmt_SumToTotalPayments_IsConsistent()
    {
        // CUMPRINC + CUMIPMT should equal total payments for the period
        var rate = 0.06 / 12;
        var nper = 360.0;
        var pv = 100000.0;
        var startPeriod = 1.0;
        var endPeriod = 12.0;
        var type = 0.0;

        var cumprincFunc = CumprincFunction.Instance;
        var cumprincArgs = new[]
        {
            CellValue.FromNumber(rate),
            CellValue.FromNumber(nper),
            CellValue.FromNumber(pv),
            CellValue.FromNumber(startPeriod),
            CellValue.FromNumber(endPeriod),
            CellValue.FromNumber(type),
        };

        var cumprincResult = cumprincFunc.Execute(null!, cumprincArgs);
        Assert.Equal(CellValueType.Number, cumprincResult.Type);

        var cumipmtFunc = CumipmtFunction.Instance;
        var cumipmtArgs = new[]
        {
            CellValue.FromNumber(rate),
            CellValue.FromNumber(nper),
            CellValue.FromNumber(pv),
            CellValue.FromNumber(startPeriod),
            CellValue.FromNumber(endPeriod),
            CellValue.FromNumber(type),
        };

        var cumipmtResult = cumipmtFunc.Execute(null!, cumipmtArgs);
        Assert.Equal(CellValueType.Number, cumipmtResult.Type);

        // Calculate total payment for 12 months
        var pmtFunc = PmtFunction.Instance;
        var pmtArgs = new[]
        {
            CellValue.FromNumber(rate),
            CellValue.FromNumber(nper),
            CellValue.FromNumber(pv),
        };

        var pmtResult = pmtFunc.Execute(null!, pmtArgs);
        var totalPayments = pmtResult.NumericValue * (endPeriod - startPeriod + 1);

        var sum = cumprincResult.NumericValue + cumipmtResult.NumericValue;
        Assert.Equal(totalPayments, sum, 2);
    }

    // FVSCHEDULE Function Tests
    [Fact]
    public void Fvschedule_VariableRates_ReturnsCorrectValue()
    {
        // FVSCHEDULE(1000, 0.09, 0.11, 0.10)
        var func = FvscheduleFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1000),   // principal
            CellValue.FromNumber(0.09),   // rate1
            CellValue.FromNumber(0.11),   // rate2
            CellValue.FromNumber(0.10),   // rate3
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        // Expected: 1000 * 1.09 * 1.11 * 1.10 = 1330.89
        Assert.Equal(1330.89, result.NumericValue, 2);
    }

    [Fact]
    public void Fvschedule_SingleRate_ReturnsCorrectValue()
    {
        var func = FvscheduleFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1000),
            CellValue.FromNumber(0.05),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1050.0, result.NumericValue, 2);
    }

    [Fact]
    public void Fvschedule_NegativeRates_ReturnsCorrectValue()
    {
        // Negative rates should reduce the value
        var func = FvscheduleFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1000),
            CellValue.FromNumber(-0.05),
            CellValue.FromNumber(0.10),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1045.0, result.NumericValue, 2); // 1000 * 0.95 * 1.10
    }

    // XNPV Function Tests
    [Fact]
    public void Xnpv_IrregularCashFlows_ReturnsCorrectValue()
    {
        // XNPV with irregular dates
        // Rate: 9%, Cash flows on different dates
        var func = XnpvFunction.Instance;

        // Using Excel date serial numbers
        // Jan 1, 2024 = 45292, Feb 1, 2024 = 45323, etc.
        var args = new[]
        {
            CellValue.FromNumber(0.09),      // rate
            CellValue.FromNumber(-10000),    // value1
            CellValue.FromNumber(45292),     // date1 (Jan 1, 2024)
            CellValue.FromNumber(2750),      // value2
            CellValue.FromNumber(45323),     // date2 (Feb 1, 2024)
            CellValue.FromNumber(4250),      // value3
            CellValue.FromNumber(45506),     // date3 (Jul 3, 2024)
            CellValue.FromNumber(3250),      // value4
            CellValue.FromNumber(45657),     // date4 (Dec 1, 2024)
            CellValue.FromNumber(2750),      // value5
            CellValue.FromNumber(45810),     // date5 (May 3, 2025)
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue > 2000); // Approximate positive NPV
    }

    [Fact]
    public void Xnpv_InvalidArgCount_ReturnsError()
    {
        var func = XnpvFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(0.09),
            CellValue.FromNumber(-10000),
            // Missing date
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#VALUE!", result.ErrorValue);
    }

    // XIRR Function Tests
    [Fact]
    public void Xirr_IrregularCashFlows_ReturnsCorrectValue()
    {
        // XIRR with irregular dates
        var func = XirrFunction.Instance;

        var args = new[]
        {
            CellValue.FromNumber(-10000),    // value1
            CellValue.FromNumber(45292),     // date1 (Jan 1, 2024)
            CellValue.FromNumber(2750),      // value2
            CellValue.FromNumber(45323),     // date2 (Feb 1, 2024)
            CellValue.FromNumber(4250),      // value3
            CellValue.FromNumber(45506),     // date3 (Jul 3, 2024)
            CellValue.FromNumber(3250),      // value4
            CellValue.FromNumber(45657),     // date4 (Dec 1, 2024)
            CellValue.FromNumber(2750),      // value5
            CellValue.FromNumber(45810),     // date5 (May 3, 2025)
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue > 0); // Should have positive return
        Assert.True(result.NumericValue < 1); // Reasonable rate
    }

    [Fact]
    public void Xirr_AllPositive_ReturnsError()
    {
        var func = XirrFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1000),
            CellValue.FromNumber(45292),
            CellValue.FromNumber(2000),
            CellValue.FromNumber(45323),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Xirr_AllNegative_ReturnsError()
    {
        var func = XirrFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-1000),
            CellValue.FromNumber(45292),
            CellValue.FromNumber(-2000),
            CellValue.FromNumber(45323),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    [Fact]
    public void Xirr_WithGuess_ReturnsCorrectValue()
    {
        // XIRR with optional guess parameter
        var func = XirrFunction.Instance;

        var args = new[]
        {
            CellValue.FromNumber(-10000),
            CellValue.FromNumber(45292),
            CellValue.FromNumber(2750),
            CellValue.FromNumber(45323),
            CellValue.FromNumber(4250),
            CellValue.FromNumber(45506),
            CellValue.FromNumber(3250),
            CellValue.FromNumber(45657),
            CellValue.FromNumber(2750),
            CellValue.FromNumber(45810),
            CellValue.FromNumber(0.15),  // guess
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue > 0);
        Assert.True(result.NumericValue < 1);
    }

    // Error Propagation Tests
    [Fact]
    public void Effect_PropagatesError()
    {
        var func = EffectFunction.Instance;
        var args = new[]
        {
            CellValue.Error("#DIV/0!"),
            CellValue.FromNumber(4),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#DIV/0!", result.ErrorValue);
    }

    [Fact]
    public void Mirr_PropagatesError()
    {
        var func = MirrFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(-10000),
            CellValue.Error("#N/A"),
            CellValue.FromNumber(3000),
            CellValue.FromNumber(0.1),
            CellValue.FromNumber(0.12),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#N/A", result.ErrorValue);
    }

    [Fact]
    public void Fvschedule_PropagatesError()
    {
        var func = FvscheduleFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1000),
            CellValue.FromNumber(0.09),
            CellValue.Error("#REF!"),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#REF!", result.ErrorValue);
    }
}
