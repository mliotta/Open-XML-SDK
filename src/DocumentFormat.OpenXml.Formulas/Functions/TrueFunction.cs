// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the TRUE function.
/// TRUE() - Returns the logical value TRUE.
/// </summary>
public sealed class TrueFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly TrueFunction Instance = new();

    private TrueFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "TRUE";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 0)
        {
            return CellValue.Error("#VALUE!");
        }

        return CellValue.FromBool(true);
    }
}
