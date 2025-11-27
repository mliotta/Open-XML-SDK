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
    // Helper to create date values
    private static FormulaResult DateValue(int year, int month, int day)
    {
        return FormulaResult.FromNumber(new DateTime(year, month, day).ToOADate());
    }

    [Fact]
    public void ACCRINT_BasicCalculation_ReturnsAccruedInterest()
    {
        var func = AccrintFunction.Instance;
        var args = new[]
        {
            DateValue(2023, 1, 1),  // issue
            DateValue(2023, 7, 1),  // first_interest
            DateValue(2023, 3, 1),  // settlement
            FormulaResult.FromNumber(0.05),  // rate
            FormulaResult.FromNumber(1000),  // par
            FormulaResult.FromNumber(2),  // frequency (semi-annual)
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.True(result.NumericValue > 0);
    }

    [Fact]
    public void ACCRINTM_BasicCalculation_ReturnsAccruedInterest()
    {
        var func = AccrintmFunction.Instance;
        var args = new[]
        {
            DateValue(2023, 1, 1),  // issue
            DateValue(2023, 12, 31),  // settlement
            FormulaResult.FromNumber(0.05),  // rate
            FormulaResult.FromNumber(1000),  // par
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.True(result.NumericValue > 0);
    }

    [Fact]
    public void DISC_BasicCalculation_ReturnsDiscountRate()
    {
        var func = DiscFunction.Instance;
        var args = new[]
        {
            DateValue(2023, 1, 1),  // settlement
            DateValue(2023, 12, 31),  // maturity
            FormulaResult.FromNumber(95),  // pr
            FormulaResult.FromNumber(100),  // redemption
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.True(result.NumericValue > 0);
    }

    [Fact]
    public void INTRATE_BasicCalculation_ReturnsInterestRate()
    {
        var func = IntrateFunction.Instance;
        var args = new[]
        {
            DateValue(2023, 1, 1),  // settlement
            DateValue(2023, 12, 31),  // maturity
            FormulaResult.FromNumber(1000),  // investment
            FormulaResult.FromNumber(1050),  // redemption
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.True(result.NumericValue > 0);
    }

    [Fact]
    public void RECEIVED_BasicCalculation_ReturnsAmountReceived()
    {
        var func = ReceivedFunction.Instance;
        var args = new[]
        {
            DateValue(2023, 1, 1),  // settlement
            DateValue(2023, 12, 31),  // maturity
            FormulaResult.FromNumber(1000),  // investment
            FormulaResult.FromNumber(0.05),  // discount
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.True(result.NumericValue > 1000);
    }

    [Fact]
    public void TBILLPRICE_BasicCalculation_ReturnsPrice()
    {
        var func = TbillpriceFunction.Instance;
        var args = new[]
        {
            DateValue(2023, 1, 1),  // settlement
            DateValue(2023, 3, 31),  // maturity (90 days)
            FormulaResult.FromNumber(0.05),  // discount
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.True(result.NumericValue < 100 && result.NumericValue > 0);
    }

    [Fact]
    public void TBILLYIELD_BasicCalculation_ReturnsYield()
    {
        var func = TbillyieldFunction.Instance;
        var args = new[]
        {
            DateValue(2023, 1, 1),  // settlement
            DateValue(2023, 3, 31),  // maturity (90 days)
            FormulaResult.FromNumber(98.75),  // pr
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.True(result.NumericValue > 0);
    }

    [Fact]
    public void VDB_BasicCalculation_ReturnsDepreciation()
    {
        var func = VdbFunction.Instance;
        var args = new[]
        {
            FormulaResult.FromNumber(10000),  // cost
            FormulaResult.FromNumber(1000),  // salvage
            FormulaResult.FromNumber(10),  // life
            FormulaResult.FromNumber(0),  // start_period
            FormulaResult.FromNumber(1),  // end_period
        };

        var result = func.Execute(null!, args);

        Assert.Equal(FormulaResultType.Number, result.Type);
        Assert.True(result.NumericValue > 0);
    }

    [Fact]
    public void ACCRINT_InvalidDates_ReturnsError()
    {
        var func = AccrintFunction.Instance;
        var args = new[]
        {
            DateValue(2023, 3, 1),  // issue (after settlement)
            DateValue(2023, 7, 1),  // first_interest
            DateValue(2023, 1, 1),  // settlement
            FormulaResult.FromNumber(0.05),  // rate
            FormulaResult.FromNumber(1000),  // par
            FormulaResult.FromNumber(2),  // frequency
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
    }
}
