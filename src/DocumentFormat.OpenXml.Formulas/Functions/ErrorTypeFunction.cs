// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ERROR.TYPE function.
/// ERROR.TYPE(error_val) - Returns a number corresponding to an error type.
/// </summary>
/// <remarks>
/// Error type codes:
/// 1 = #NULL!
/// 2 = #DIV/0!
/// 3 = #VALUE!
/// 4 = #REF!
/// 5 = #NAME?
/// 6 = #NUM!
/// 7 = #N/A
/// 8 = #GETTING_DATA
/// </remarks>
public sealed class ErrorTypeFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ErrorTypeFunction Instance = new();

    private ErrorTypeFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ERROR.TYPE";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 1)
        {
            return CellValue.Error("#VALUE!");
        }

        // ERROR.TYPE returns #N/A if the value is not an error
        if (!args[0].IsError)
        {
            return CellValue.Error("#N/A");
        }

        var errorValue = args[0].ErrorValue;
        var errorType = errorValue switch
        {
            "#NULL!" => 1,
            "#DIV/0!" => 2,
            "#VALUE!" => 3,
            "#REF!" => 4,
            "#NAME?" => 5,
            "#NUM!" => 6,
            "#N/A" => 7,
            "#GETTING_DATA" => 8,
            _ => 7, // Default to #N/A if unknown error type
        };

        return CellValue.FromNumber(errorType);
    }
}
