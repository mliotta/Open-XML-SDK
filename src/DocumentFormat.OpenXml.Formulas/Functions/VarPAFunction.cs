// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the VARPA function.
/// VARPA(value1, [value2], ...) - Variance (population) including text and logical values.
/// Text evaluates as 0, TRUE as 1, FALSE as 0, empty values are ignored.
/// </summary>
public sealed class VarPAFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly VarPAFunction Instance = new();

    private VarPAFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "VARPA";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        var values = new List<double>();

        foreach (var arg in args)
        {
            if (arg.IsError)
            {
                return arg; // Propagate errors
            }

            if (arg.Type == CellValueType.Number)
            {
                values.Add(arg.NumericValue);
            }
            else if (arg.Type == CellValueType.Boolean)
            {
                values.Add(arg.BoolValue ? 1.0 : 0.0);
            }
            else if (arg.Type == CellValueType.Text)
            {
                // Text values count as 0
                values.Add(0.0);
            }
            // Empty values are ignored
        }

        if (values.Count == 0)
        {
            return CellValue.Error("#DIV/0!");
        }

        // Calculate mean
        var mean = values.Average();

        // Calculate variance (population)
        var sumSquaredDiffs = values.Sum(v => System.Math.Pow(v - mean, 2));
        var variance = sumSquaredDiffs / values.Count;

        return CellValue.FromNumber(variance);
    }
}
