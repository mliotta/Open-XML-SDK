// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests.Functions;

/// <summary>
/// Tests for securities/bond pricing functions.
/// </summary>
public class SecuritiesFunctionTests
{
    // DOLLARDE Function Tests
    [Fact]
    public void Dollarde_ConvertFractionalToDecimal_ReturnsCorrectValue()
    {
        // DOLLARDE(1.02, 16) - converts 1 and 2/16 to decimal
        var func = DollardeFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1.02),
            CellValue.FromNumber(16),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.125, result.NumericValue, 3); // 1 + 2/16 = 1.125
    }

    [Fact]
    public void Dollarde_ConvertEights_ReturnsCorrectValue()
    {
        // DOLLARDE(1.1, 32) - converts 1 and 1/32
        var func = DollardeFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1.1),
            CellValue.FromNumber(32),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.03125, result.NumericValue, 5);
    }

    [Fact]
    public void Dollarde_InvalidFraction_ReturnsError()
    {
        var func = DollardeFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1.02),
            CellValue.FromNumber(0),  // invalid
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    // DOLLARFR Function Tests
    [Fact]
    public void Dollarfr_ConvertDecimalToFractional_ReturnsCorrectValue()
    {
        // DOLLARFR(1.125, 16) - converts 1.125 to 1 and 2/16
        var func = DollarfrFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1.125),
            CellValue.FromNumber(16),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.02, result.NumericValue, 2);
    }

    [Fact]
    public void Dollarfr_ConvertThirtySeconds_ReturnsCorrectValue()
    {
        // DOLLARFR(1.03125, 32)
        var func = DollarfrFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(1.03125),
            CellValue.FromNumber(32),
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(1.01, result.NumericValue, 2);
    }

    // COUPNCD Function Tests
    [Fact]
    public void Coupncd_SemiAnnualBond_ReturnsNextCouponDate()
    {
        // COUPNCD(settlement, maturity, frequency=2)
        var func = CoupncdFunction.Instance;
        var settlement = new DateTime(2023, 1, 25).ToOADate();
        var maturity = new DateTime(2030, 11, 15).ToOADate();

        var args = new[]
        {
            CellValue.FromNumber(settlement),
            CellValue.FromNumber(maturity),
            CellValue.FromNumber(2),  // Semi-annual
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        var nextCoupon = DateTime.FromOADate(result.NumericValue);
        Assert.Equal(5, nextCoupon.Month); // May 15
        Assert.Equal(15, nextCoupon.Day);
    }

    [Fact]
    public void Coupncd_InvalidFrequency_ReturnsError()
    {
        var func = CoupncdFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(new DateTime(2023, 1, 25).ToOADate()),
            CellValue.FromNumber(new DateTime(2030, 11, 15).ToOADate()),
            CellValue.FromNumber(3),  // Invalid frequency
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    // COUPPCD Function Tests
    [Fact]
    public void Couppcd_SemiAnnualBond_ReturnsPreviousCouponDate()
    {
        var func = CouppcdFunction.Instance;
        var settlement = new DateTime(2023, 1, 25).ToOADate();
        var maturity = new DateTime(2030, 11, 15).ToOADate();

        var args = new[]
        {
            CellValue.FromNumber(settlement),
            CellValue.FromNumber(maturity),
            CellValue.FromNumber(2),  // Semi-annual
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        var prevCoupon = DateTime.FromOADate(result.NumericValue);
        Assert.Equal(11, prevCoupon.Month); // November 15
        Assert.Equal(15, prevCoupon.Day);
        Assert.Equal(2022, prevCoupon.Year);
    }

    // COUPNUM Function Tests
    [Fact]
    public void Coupnum_SemiAnnualBond_ReturnsCorrectCount()
    {
        var func = CoupnumFunction.Instance;
        var settlement = new DateTime(2023, 1, 25).ToOADate();
        var maturity = new DateTime(2030, 11, 15).ToOADate();

        var args = new[]
        {
            CellValue.FromNumber(settlement),
            CellValue.FromNumber(maturity),
            CellValue.FromNumber(2),  // Semi-annual
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(16, result.NumericValue); // ~7.8 years * 2 = 16 coupons
    }

    [Fact]
    public void Coupnum_QuarterlyBond_ReturnsCorrectCount()
    {
        var func = CoupnumFunction.Instance;
        var settlement = new DateTime(2023, 1, 25).ToOADate();
        var maturity = new DateTime(2024, 11, 15).ToOADate();

        var args = new[]
        {
            CellValue.FromNumber(settlement),
            CellValue.FromNumber(maturity),
            CellValue.FromNumber(4),  // Quarterly
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue >= 7 && result.NumericValue <= 8);
    }

    // COUPDAYBS Function Tests
    [Fact]
    public void Coupdaybs_ReturnsPositiveValue()
    {
        var func = CoupdaybsFunction.Instance;
        var settlement = new DateTime(2023, 1, 25).ToOADate();
        var maturity = new DateTime(2030, 11, 15).ToOADate();

        var args = new[]
        {
            CellValue.FromNumber(settlement),
            CellValue.FromNumber(maturity),
            CellValue.FromNumber(2),  // Semi-annual
            CellValue.FromNumber(0),  // 30/360 US
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue > 0);
    }

    // COUPDAYS Function Tests
    [Fact]
    public void Coupdays_SemiAnnualWith30360_Returns180()
    {
        var func = CoupdaysFunction.Instance;
        var settlement = new DateTime(2023, 1, 25).ToOADate();
        var maturity = new DateTime(2030, 11, 15).ToOADate();

        var args = new[]
        {
            CellValue.FromNumber(settlement),
            CellValue.FromNumber(maturity),
            CellValue.FromNumber(2),  // Semi-annual
            CellValue.FromNumber(0),  // 30/360 US
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(180, result.NumericValue, 0); // 360/2 = 180
    }

    // COUPDAYSNC Function Tests
    [Fact]
    public void Coupdaysnc_ReturnsPositiveValue()
    {
        var func = CoupdaysncFunction.Instance;
        var settlement = new DateTime(2023, 1, 25).ToOADate();
        var maturity = new DateTime(2030, 11, 15).ToOADate();

        var args = new[]
        {
            CellValue.FromNumber(settlement),
            CellValue.FromNumber(maturity),
            CellValue.FromNumber(2),  // Semi-annual
            CellValue.FromNumber(0),  // 30/360 US
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue > 0);
    }

    // PRICE Function Tests
    [Fact]
    public void Price_SemiAnnualBond_ReturnsCorrectPrice()
    {
        var func = PriceFunction.Instance;
        var settlement = new DateTime(2008, 2, 15).ToOADate();
        var maturity = new DateTime(2017, 11, 15).ToOADate();

        var args = new[]
        {
            CellValue.FromNumber(settlement),
            CellValue.FromNumber(maturity),
            CellValue.FromNumber(0.0575),  // 5.75% coupon
            CellValue.FromNumber(0.065),   // 6.5% yield
            CellValue.FromNumber(100),     // Redemption value
            CellValue.FromNumber(2),       // Semi-annual
            CellValue.FromNumber(0),       // 30/360 basis
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue > 90 && result.NumericValue < 100); // Discount bond
    }

    [Fact]
    public void Price_InvalidInputs_ReturnsError()
    {
        var func = PriceFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(new DateTime(2023, 1, 1).ToOADate()),
            CellValue.FromNumber(new DateTime(2022, 1, 1).ToOADate()), // maturity < settlement
            CellValue.FromNumber(0.05),
            CellValue.FromNumber(0.06),
            CellValue.FromNumber(100),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    // PRICEDISC Function Tests
    [Fact]
    public void Pricedisc_DiscountedSecurity_ReturnsCorrectPrice()
    {
        var func = PricediscFunction.Instance;
        var settlement = new DateTime(2007, 2, 16).ToOADate();
        var maturity = new DateTime(2007, 3, 1).ToOADate();

        var args = new[]
        {
            CellValue.FromNumber(settlement),
            CellValue.FromNumber(maturity),
            CellValue.FromNumber(0.0525),  // 5.25% discount rate
            CellValue.FromNumber(100),     // Redemption value
            CellValue.FromNumber(2),       // Actual/360
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue > 99 && result.NumericValue < 100);
    }

    // PRICEMAT Function Tests
    [Fact]
    public void Pricemat_MaturityPayment_ReturnsCorrectPrice()
    {
        var func = PricematFunction.Instance;
        var settlement = new DateTime(2008, 2, 15).ToOADate();
        var maturity = new DateTime(2008, 4, 13).ToOADate();
        var issue = new DateTime(2007, 11, 11).ToOADate();

        var args = new[]
        {
            CellValue.FromNumber(settlement),
            CellValue.FromNumber(maturity),
            CellValue.FromNumber(issue),
            CellValue.FromNumber(0.061),   // 6.1% interest rate
            CellValue.FromNumber(0.061),   // 6.1% yield
            CellValue.FromNumber(0),       // 30/360 basis
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue > 95 && result.NumericValue < 105);
    }

    // YIELDDISC Function Tests
    [Fact]
    public void Yielddisc_DiscountedSecurity_ReturnsCorrectYield()
    {
        var func = YielddiscFunction.Instance;
        var settlement = new DateTime(2007, 2, 16).ToOADate();
        var maturity = new DateTime(2007, 3, 1).ToOADate();

        var args = new[]
        {
            CellValue.FromNumber(settlement),
            CellValue.FromNumber(maturity),
            CellValue.FromNumber(99.795),  // Price
            CellValue.FromNumber(100),     // Redemption
            CellValue.FromNumber(2),       // Actual/360
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue > 0);
    }

    // YIELDMAT Function Tests
    [Fact]
    public void Yieldmat_MaturityPayment_ReturnsCorrectYield()
    {
        var func = YieldmatFunction.Instance;
        var settlement = new DateTime(2008, 3, 15).ToOADate();
        var maturity = new DateTime(2008, 11, 3).ToOADate();
        var issue = new DateTime(2007, 11, 8).ToOADate();

        var args = new[]
        {
            CellValue.FromNumber(settlement),
            CellValue.FromNumber(maturity),
            CellValue.FromNumber(issue),
            CellValue.FromNumber(0.0625),  // 6.25% interest rate
            CellValue.FromNumber(100.0123),// Price
            CellValue.FromNumber(0),       // 30/360 basis
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue > 0);
    }

    // DURATION Function Tests
    [Fact]
    public void Duration_SemiAnnualBond_ReturnsCorrectDuration()
    {
        var func = DurationFunction.Instance;
        var settlement = new DateTime(2008, 1, 1).ToOADate();
        var maturity = new DateTime(2016, 1, 1).ToOADate();

        var args = new[]
        {
            CellValue.FromNumber(settlement),
            CellValue.FromNumber(maturity),
            CellValue.FromNumber(0.08),    // 8% coupon
            CellValue.FromNumber(0.09),    // 9% yield
            CellValue.FromNumber(2),       // Semi-annual
            CellValue.FromNumber(1),       // Actual/actual
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue > 5 && result.NumericValue < 8); // Reasonable duration range
    }

    [Fact]
    public void Duration_InvalidInputs_ReturnsError()
    {
        var func = DurationFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(new DateTime(2023, 1, 1).ToOADate()),
            CellValue.FromNumber(new DateTime(2022, 1, 1).ToOADate()), // maturity < settlement
            CellValue.FromNumber(0.05),
            CellValue.FromNumber(0.06),
            CellValue.FromNumber(2),
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
        Assert.Equal("#NUM!", result.ErrorValue);
    }

    // MDURATION Function Tests
    [Fact]
    public void Mduration_SemiAnnualBond_ReturnsCorrectDuration()
    {
        var func = MdurationFunction.Instance;
        var settlement = new DateTime(2008, 1, 1).ToOADate();
        var maturity = new DateTime(2016, 1, 1).ToOADate();

        var args = new[]
        {
            CellValue.FromNumber(settlement),
            CellValue.FromNumber(maturity),
            CellValue.FromNumber(0.08),    // 8% coupon
            CellValue.FromNumber(0.09),    // 9% yield
            CellValue.FromNumber(2),       // Semi-annual
            CellValue.FromNumber(1),       // Actual/actual
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue > 5 && result.NumericValue < 8);
    }

    [Fact]
    public void Mduration_LessThanMacaulay_IsConsistent()
    {
        // Modified duration should be less than Macaulay duration
        var settlement = new DateTime(2008, 1, 1).ToOADate();
        var maturity = new DateTime(2016, 1, 1).ToOADate();

        var durationArgs = new[]
        {
            CellValue.FromNumber(settlement),
            CellValue.FromNumber(maturity),
            CellValue.FromNumber(0.08),
            CellValue.FromNumber(0.09),
            CellValue.FromNumber(2),
            CellValue.FromNumber(1),
        };

        var macaulay = DurationFunction.Instance.Execute(null!, durationArgs);
        var modified = MdurationFunction.Instance.Execute(null!, durationArgs);

        Assert.Equal(CellValueType.Number, macaulay.Type);
        Assert.Equal(CellValueType.Number, modified.Type);
        Assert.True(modified.NumericValue < macaulay.NumericValue);
    }

    // Error Propagation Tests
    [Fact]
    public void SecuritiesFunctions_PropagateErrors()
    {
        var functions = new IFunctionImplementation[]
        {
            DollardeFunction.Instance,
            DollarfrFunction.Instance,
            CoupncdFunction.Instance,
            CouppcdFunction.Instance,
            CoupnumFunction.Instance,
            CoupdaybsFunction.Instance,
            CoupdaysFunction.Instance,
            CoupdaysncFunction.Instance,
        };

        foreach (var func in functions)
        {
            var args = new[] { CellValue.Error("#DIV/0!"), CellValue.FromNumber(1) };
            var result = func.Execute(null!, args);

            Assert.True(result.IsError);
            Assert.Equal("#DIV/0!", result.ErrorValue);
        }
    }
}
