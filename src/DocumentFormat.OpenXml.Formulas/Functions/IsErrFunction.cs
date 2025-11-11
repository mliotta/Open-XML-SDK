// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ISERR function.
/// ISERR(value) - TRUE if value is any error except #N/A.
/// </summary>
public sealed class IsErrFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly IsErrFunction Instance = new();

    private IsErrFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ISERR";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 1)
        {
            return CellValue.Error("#VALUE!");
        }

        // Note: Errors are NOT propagated for IS* functions
        // Check if the value is an error but NOT #N/A
        var isErr = args[0].IsError && args[0].ErrorValue != "#N/A";
        return CellValue.FromBool(isErr);
    }
}
