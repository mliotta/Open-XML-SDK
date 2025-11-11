// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ISNA function.
/// ISNA(value) - TRUE if value is #N/A error.
/// </summary>
public sealed class IsNaFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly IsNaFunction Instance = new();

    private IsNaFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ISNA";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 1)
        {
            return CellValue.Error("#VALUE!");
        }

        // Note: Errors are NOT propagated for IS* functions
        // Check if the value is specifically the #N/A error
        var isNa = args[0].IsError && args[0].ErrorValue == "#N/A";
        return CellValue.FromBool(isNa);
    }
}
