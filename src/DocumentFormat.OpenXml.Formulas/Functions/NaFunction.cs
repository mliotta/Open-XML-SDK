// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the NA function.
/// NA() - returns the #N/A error value.
/// </summary>
public sealed class NaFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly NaFunction Instance = new();

    private NaFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "NA";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // NA() takes no arguments and always returns #N/A
        return CellValue.Error("#N/A");
    }
}
