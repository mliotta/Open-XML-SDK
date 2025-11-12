// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the STDEV.P function.
/// STDEV.P(number1, [number2], ...) - standard deviation (population) - Excel 2010+ compatibility function.
/// This is the same as STDEVP.
/// </summary>
public sealed class StDevPFunction2 : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly StDevPFunction2 Instance = new();

    private StDevPFunction2()
    {
    }

    /// <inheritdoc/>
    public string Name => "STDEV.P";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // Delegate to STDEVP implementation
        return StDevPFunction.Instance.Execute(context, args);
    }
}
