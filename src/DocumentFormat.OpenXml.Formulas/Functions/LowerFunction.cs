// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the LOWER function.
/// LOWER(text) - converts to lowercase.
/// </summary>
public sealed class LowerFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly LowerFunction Instance = new();

    private LowerFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "LOWER";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 1)
        {
            return CellValue.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        return CellValue.FromString(args[0].StringValue.ToLowerInvariant());
    }
}
