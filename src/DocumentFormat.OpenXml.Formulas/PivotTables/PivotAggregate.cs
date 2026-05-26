// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace DocumentFormat.OpenXml.Features.PivotTables;

/// <summary>
/// Identifies the consolidation function applied to a pivot table value field.
/// </summary>
/// <remarks>
/// Each member maps to a <see cref="DocumentFormat.OpenXml.Spreadsheet.DataConsolidateFunctionValues"/>
/// for the emitted definition and to an existing built-in function in
/// <see cref="DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions.FunctionRegistry"/> used to
/// precompute the displayed values.
/// </remarks>
public enum PivotAggregate
{
    /// <summary>Sum of the numeric values (<c>SUM</c>).</summary>
    Sum,

    /// <summary>Count of non-empty values (<c>COUNTA</c>).</summary>
    Count,

    /// <summary>Count of numeric values only (<c>COUNT</c>).</summary>
    CountNumbers,

    /// <summary>Arithmetic mean of the numeric values (<c>AVERAGE</c>).</summary>
    Average,

    /// <summary>Largest numeric value (<c>MAX</c>).</summary>
    Max,

    /// <summary>Smallest numeric value (<c>MIN</c>).</summary>
    Min,

    /// <summary>Product of the numeric values (<c>PRODUCT</c>).</summary>
    Product,

    /// <summary>Sample standard deviation (<c>STDEV</c>).</summary>
    StdDev,

    /// <summary>Population standard deviation (<c>STDEVP</c>).</summary>
    StdDevP,

    /// <summary>Sample variance (<c>VAR</c>).</summary>
    Var,

    /// <summary>Population variance (<c>VARP</c>).</summary>
    VarP,
}
