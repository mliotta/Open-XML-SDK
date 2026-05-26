// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;

using DocumentFormat.OpenXml.Spreadsheet;

namespace DocumentFormat.OpenXml.Features.PivotTables;

/// <summary>
/// Maps <see cref="PivotAggregate"/> values to the OOXML subtotal enumeration and to the
/// built-in function name used to precompute results.
/// </summary>
internal static class PivotAggregateMap
{
    /// <summary>
    /// Gets the name of the built-in function (in the formula <c>FunctionRegistry</c>) that
    /// computes the supplied aggregate.
    /// </summary>
    /// <param name="aggregate">The aggregate.</param>
    /// <returns>The uppercase function name.</returns>
    public static string FunctionName(PivotAggregate aggregate) => aggregate switch
    {
        PivotAggregate.Sum => "SUM",
        PivotAggregate.Count => "COUNTA",
        PivotAggregate.CountNumbers => "COUNT",
        PivotAggregate.Average => "AVERAGE",
        PivotAggregate.Max => "MAX",
        PivotAggregate.Min => "MIN",
        PivotAggregate.Product => "PRODUCT",
        PivotAggregate.StdDev => "STDEV",
        PivotAggregate.StdDevP => "STDEVP",
        PivotAggregate.Var => "VAR",
        PivotAggregate.VarP => "VARP",
        _ => throw new ArgumentOutOfRangeException(nameof(aggregate), aggregate, "Unsupported pivot aggregate."),
    };

    /// <summary>
    /// Gets the OOXML <see cref="DataConsolidateFunctionValues"/> for the supplied aggregate.
    /// </summary>
    /// <param name="aggregate">The aggregate.</param>
    /// <returns>The matching consolidation function value.</returns>
    public static DataConsolidateFunctionValues Subtotal(PivotAggregate aggregate) => aggregate switch
    {
        PivotAggregate.Sum => DataConsolidateFunctionValues.Sum,
        PivotAggregate.Count => DataConsolidateFunctionValues.Count,
        PivotAggregate.CountNumbers => DataConsolidateFunctionValues.CountNumbers,
        PivotAggregate.Average => DataConsolidateFunctionValues.Average,
        PivotAggregate.Max => DataConsolidateFunctionValues.Maximum,
        PivotAggregate.Min => DataConsolidateFunctionValues.Minimum,
        PivotAggregate.Product => DataConsolidateFunctionValues.Product,
        PivotAggregate.StdDev => DataConsolidateFunctionValues.StandardDeviation,
        PivotAggregate.StdDevP => DataConsolidateFunctionValues.StandardDeviationP,
        PivotAggregate.Var => DataConsolidateFunctionValues.Variance,
        PivotAggregate.VarP => DataConsolidateFunctionValues.VarianceP,
        _ => throw new ArgumentOutOfRangeException(nameof(aggregate), aggregate, "Unsupported pivot aggregate."),
    };

    /// <summary>
    /// Gets the OOXML <see cref="ShowDataAsValues"/> for a display mode, or null for
    /// <see cref="PivotShowAs.Normal"/>.
    /// </summary>
    /// <param name="showAs">The display mode.</param>
    /// <returns>The matching value, or null.</returns>
    public static ShowDataAsValues? ShowDataAs(PivotShowAs showAs) => showAs switch
    {
        PivotShowAs.PercentOfTotal => ShowDataAsValues.PercentOfTotal,
        PivotShowAs.PercentOfColumn => ShowDataAsValues.PercentOfColumn,
        PivotShowAs.PercentOfRow => ShowDataAsValues.PercentOfRaw,
        _ => null,
    };

    /// <summary>
    /// Gets the default display caption Excel uses for an aggregate over a field
    /// (e.g. <c>"Sum of Sales"</c>).
    /// </summary>
    /// <param name="aggregate">The aggregate.</param>
    /// <param name="fieldName">The source field name.</param>
    /// <returns>The default caption.</returns>
    public static string DisplayName(PivotAggregate aggregate, string fieldName)
    {
        var verb = aggregate switch
        {
            PivotAggregate.Sum => "Sum",
            PivotAggregate.Count => "Count",
            PivotAggregate.CountNumbers => "Count",
            PivotAggregate.Average => "Average",
            PivotAggregate.Max => "Max",
            PivotAggregate.Min => "Min",
            PivotAggregate.Product => "Product",
            PivotAggregate.StdDev => "StdDev",
            PivotAggregate.StdDevP => "StdDevp",
            PivotAggregate.Var => "Var",
            PivotAggregate.VarP => "Varp",
            _ => "Sum",
        };

        return $"{verb} of {fieldName}";
    }
}
