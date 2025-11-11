// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the EXACT function.
/// EXACT(text1, text2) - case-sensitive text comparison (returns TRUE/FALSE).
/// </summary>
public sealed class ExactFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ExactFunction Instance = new();

    private ExactFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "EXACT";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 2)
        {
            return CellValue.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[1].IsError)
        {
            return args[1];
        }

        var text1 = args[0].StringValue;
        var text2 = args[1].StringValue;

        // Case-sensitive comparison
        var isEqual = string.Equals(text1, text2, StringComparison.Ordinal);

        return CellValue.FromBool(isEqual);
    }
}
