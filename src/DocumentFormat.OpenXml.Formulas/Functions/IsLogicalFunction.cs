// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ISLOGICAL function.
/// ISLOGICAL(value) - TRUE if value is boolean.
/// </summary>
public sealed class IsLogicalFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly IsLogicalFunction Instance = new();

    private IsLogicalFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ISLOGICAL";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 1)
        {
            return CellValue.Error("#VALUE!");
        }

        // Note: Errors are NOT propagated for IS* functions
        var isLogical = args[0].Type == CellValueType.Boolean;
        return CellValue.FromBool(isLogical);
    }
}
