// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the AGGREGATE function.
/// AGGREGATE(function_num, options, ref1, [ref2], ...) - Returns an aggregate using specified function.
/// Function codes: 1=AVERAGE, 2=COUNT, 3=COUNTA, 4=MAX, 5=MIN, 6=PRODUCT, 7=STDEV.S, 8=STDEV.P,
/// 9=SUM, 10=VAR.S, 11=VAR.P, 12=MEDIAN, 13=MODE.SNL, 14=LARGE, 15=SMALL, 16=PERCENTILE.INC,
/// 17=QUARTILE.INC, 18=PERCENTILE.EXC, 19=QUARTILE.EXC
/// Options: 0=ignore nested SUBTOTAL/AGGREGATE, 1=ignore hidden rows, 2=ignore error values,
/// 3=ignore hidden rows and errors, 4=ignore nothing, 5=ignore hidden rows, 6=ignore error values,
/// 7=ignore hidden rows and errors
/// </summary>
public sealed class AggregateFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly AggregateFunction Instance = new();

    private AggregateFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "AGGREGATE";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 3)
        {
            return CellValue.Error("#VALUE!");
        }

        // First argument is the function number
        var functionNumArg = args[0];
        if (functionNumArg.IsError)
        {
            return functionNumArg;
        }

        if (functionNumArg.Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var functionNum = (int)functionNumArg.NumericValue;

        // Second argument is the options
        var optionsArg = args[1];
        if (optionsArg.IsError)
        {
            return optionsArg;
        }

        if (optionsArg.Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var options = (int)optionsArg.NumericValue;
        if (options < 0 || options > 7)
        {
            return CellValue.Error("#VALUE!");
        }

        // Get the data arguments (skip function number and options)
        var dataArgs = args.Skip(2).ToArray();

        // Filter data based on options
        // Options: 0-3 and 5-7 can ignore errors (options 2, 3, 6, 7)
        var ignoreErrors = options == 2 || options == 3 || options == 6 || options == 7;

        if (ignoreErrors)
        {
            // Filter out error values
            dataArgs = dataArgs.Where(arg => !arg.IsError).ToArray();
        }

        // Note: Currently we don't have access to hidden row/column information
        // or nested SUBTOTAL/AGGREGATE detection, so we can't fully implement
        // options 1, 3, 5, 7 (ignore hidden) or 0 (ignore nested).

        // For functions that require additional arguments (LARGE, SMALL, PERCENTILE, QUARTILE),
        // the last argument is the k value
        var requiresK = functionNum >= 14 && functionNum <= 19;
        CellValue[]? valueArgs = null;
        CellValue? kArg = null;

        if (requiresK)
        {
            if (dataArgs.Length < 2)
            {
                return CellValue.Error("#VALUE!");
            }
            kArg = dataArgs[dataArgs.Length - 1];
            valueArgs = dataArgs.Take(dataArgs.Length - 1).ToArray();
        }
        else
        {
            valueArgs = dataArgs;
        }


        return functionNum switch
        {
            1 => AverageFunction.Instance.Execute(context, valueArgs),
            2 => CountFunction.Instance.Execute(context, valueArgs),
            3 => CountAFunction.Instance.Execute(context, valueArgs),
            4 => MaxFunction.Instance.Execute(context, valueArgs),
            5 => MinFunction.Instance.Execute(context, valueArgs),
            6 => ProductFunction.Instance.Execute(context, valueArgs),
            7 => StDevFunction.Instance.Execute(context, valueArgs),
            8 => StDevPFunction.Instance.Execute(context, valueArgs),
            9 => SumFunction.Instance.Execute(context, valueArgs),
            10 => VarFunction.Instance.Execute(context, valueArgs),
            11 => VarPFunction.Instance.Execute(context, valueArgs),
            12 => MedianFunction.Instance.Execute(context, valueArgs),
            13 => ModeFunction.Instance.Execute(context, valueArgs),
            14 => LargeFunction.Instance.Execute(context, new[] { valueArgs[0], kArg!.Value }),
            15 => SmallFunction.Instance.Execute(context, new[] { valueArgs[0], kArg!.Value }),
            16 => PercentileFunction.Instance.Execute(context, new[] { valueArgs[0], kArg!.Value }),
            17 => QuartileFunction.Instance.Execute(context, new[] { valueArgs[0], kArg!.Value }),
            18 => PercentileFunction.Instance.Execute(context, new[] { valueArgs[0], kArg!.Value }), // EXC variant not implemented, using INC
            19 => QuartileFunction.Instance.Execute(context, new[] { valueArgs[0], kArg!.Value }), // EXC variant not implemented, using INC
            _ => CellValue.Error("#VALUE!")
        };
    }
}
