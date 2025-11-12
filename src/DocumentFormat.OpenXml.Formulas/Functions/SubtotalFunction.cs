// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the SUBTOTAL function.
/// SUBTOTAL(function_num, ref1, [ref2], ...) - Returns a subtotal using specified function.
/// Function codes 1-11 include hidden values, 101-111 ignore hidden values.
/// 1/101=AVERAGE, 2/102=COUNT, 3/103=COUNTA, 4/104=MAX, 5/105=MIN,
/// 6/106=PRODUCT, 7/107=STDEV, 8/108=STDEVP, 9/109=SUM, 10/110=VAR, 11/111=VARP
/// </summary>
public sealed class SubtotalFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly SubtotalFunction Instance = new();

    private SubtotalFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "SUBTOTAL";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 2)
        {
            return CellValue.Error("#VALUE!");
        }

        // First argument is the function number
        var firstArg = args[0];
        if (firstArg.IsError)
        {
            return firstArg;
        }

        if (firstArg.Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var functionNum = (int)firstArg.NumericValue;
        var ignoreHidden = functionNum > 100;
        var baseFunction = ignoreHidden ? functionNum - 100 : functionNum;

        // Get the data arguments (skip the function number)
        var dataArgs = args.Skip(1).ToArray();

        // Note: Currently we don't have access to hidden row/column information
        // so we treat all values the same way. This would need context enhancement
        // to properly filter hidden values when functionNum > 100.

        return baseFunction switch
        {
            1 => AverageFunction.Instance.Execute(context, dataArgs),
            2 => CountFunction.Instance.Execute(context, dataArgs),
            3 => CountAFunction.Instance.Execute(context, dataArgs),
            4 => MaxFunction.Instance.Execute(context, dataArgs),
            5 => MinFunction.Instance.Execute(context, dataArgs),
            6 => ProductFunction.Instance.Execute(context, dataArgs),
            7 => StDevFunction.Instance.Execute(context, dataArgs),
            8 => StDevPFunction.Instance.Execute(context, dataArgs),
            9 => SumFunction.Instance.Execute(context, dataArgs),
            10 => VarFunction.Instance.Execute(context, dataArgs),
            11 => VarPFunction.Instance.Execute(context, dataArgs),
            _ => CellValue.Error("#VALUE!")
        };
    }
}
