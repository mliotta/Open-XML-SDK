// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Text;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the CONCATENATE function.
/// CONCATENATE(text1, [text2], ...) - joins text strings.
/// </summary>
public sealed class ConcatenateFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ConcatenateFunction Instance = new();

    private ConcatenateFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "CONCATENATE";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        var result = new StringBuilder();

        foreach (var arg in args)
        {
            if (arg.IsError)
            {
                return arg; // Propagate errors
            }

            result.Append(arg.StringValue);
        }

        return CellValue.FromString(result.ToString());
    }
}
