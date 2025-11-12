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
    private static CellValue DateValue(int year, int month, int day)
    {
        return CellValue.FromNumber(new DateTime(year, month, day).ToOADate());
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
            CellValue.FromNumber(0.05),  // rate
            CellValue.FromNumber(1000),  // par
            CellValue.FromNumber(2),  // frequency (semi-annual)
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
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
            CellValue.FromNumber(0.05),  // rate
            CellValue.FromNumber(1000),  // par
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
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
            CellValue.FromNumber(95),  // pr
            CellValue.FromNumber(100),  // redemption
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
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
            CellValue.FromNumber(1000),  // investment
            CellValue.FromNumber(1050),  // redemption
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
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
            CellValue.FromNumber(1000),  // investment
            CellValue.FromNumber(0.05),  // discount
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
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
            CellValue.FromNumber(0.05),  // discount
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
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
            CellValue.FromNumber(98.75),  // pr
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
        Assert.True(result.NumericValue > 0);
    }

    [Fact]
    public void VDB_BasicCalculation_ReturnsDepreciation()
    {
        var func = VdbFunction.Instance;
        var args = new[]
        {
            CellValue.FromNumber(10000),  // cost
            CellValue.FromNumber(1000),  // salvage
            CellValue.FromNumber(10),  // life
            CellValue.FromNumber(0),  // start_period
            CellValue.FromNumber(1),  // end_period
        };

        var result = func.Execute(null!, args);

        Assert.Equal(CellValueType.Number, result.Type);
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
            CellValue.FromNumber(0.05),  // rate
            CellValue.FromNumber(1000),  // par
            CellValue.FromNumber(2),  // frequency
        };

        var result = func.Execute(null!, args);

        Assert.True(result.IsError);
    }
}
