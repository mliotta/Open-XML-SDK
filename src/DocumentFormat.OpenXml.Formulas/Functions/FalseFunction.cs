// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the FALSE function.
/// FALSE() - Returns the logical value FALSE.
/// </summary>
public sealed class FalseFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly FalseFunction Instance = new();

    private FalseFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "FALSE";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 0)
        {
            return CellValue.Error("#VALUE!");
        }

        return CellValue.FromBool(false);
    }
}
