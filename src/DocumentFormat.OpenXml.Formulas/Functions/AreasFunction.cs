// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the AREAS function.
/// AREAS(reference) - returns the number of areas in a reference.
/// An area is a range of contiguous cells or a single cell.
/// </summary>
public sealed class AreasFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly AreasFunction Instance = new();

    private AreasFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "AREAS";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 1)
        {
            return CellValue.Error("#VALUE!");
        }

        var arg = args[0];

        if (arg.IsError)
        {
            return arg;
        }

        // For a simple implementation, a single reference counts as 1 area
        // Multiple ranges separated by commas would be multiple areas
        // Since we receive a single argument, we count it as 1 area
        // Full implementation would require parsing range references with commas

        return CellValue.FromNumber(1);
    }
}
