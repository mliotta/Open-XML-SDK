// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the BESSELI function.
/// BESSELI(x, n) - returns the modified Bessel function In(x).
/// </summary>
public sealed class BesseliFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly BesseliFunction Instance = new();

    private BesseliFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "BESSELI";

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

        if (args[0].Type != CellValueType.Number || args[1].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var x = args[0].NumericValue;
        var n = (int)System.Math.Floor(args[1].NumericValue);

        if (n < 0)
        {
            return CellValue.Error("#NUM!");
        }

        try
        {
            var result = BesselI(x, n);
            if (double.IsNaN(result) || double.IsInfinity(result))
            {
                return CellValue.Error("#NUM!");
            }

            return CellValue.FromNumber(result);
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }

    /// <summary>
    /// Computes the modified Bessel function of the first kind In(x).
    /// Uses series expansion for small x and asymptotic expansion for large x.
    /// </summary>
    private static double BesselI(double x, int n)
    {
        const int maxIterations = 100;
        const double epsilon = 1e-10;

        if (System.Math.Abs(x) < 8.0)
        {
            // Series expansion for small x
            double term = 1.0;
            double sum = 1.0;
            double xHalf = x / 2.0;
            double xHalfSquared = xHalf * xHalf;

            // Calculate (x/2)^n / n!
            for (int i = 1; i <= n; i++)
            {
                term *= xHalf / i;
            }

            sum = term;

            // Series: I_n(x) = (x/2)^n * sum_{k=0}^inf [(x/2)^{2k} / (k! * (n+k)!)]
            for (int k = 1; k < maxIterations; k++)
            {
                term *= xHalfSquared / (k * (n + k));
                sum += term;

                if (System.Math.Abs(term) < epsilon * System.Math.Abs(sum))
                {
                    break;
                }
            }

            return sum;
        }
        else
        {
            // Asymptotic expansion for large x
            double result = System.Math.Exp(System.Math.Abs(x)) / System.Math.Sqrt(2.0 * System.Math.PI * System.Math.Abs(x));

            // Apply correction factor for order n
            double correction = 1.0;
            double mu = 4.0 * n * n;
            double term = 1.0;

            for (int k = 1; k < 10; k++)
            {
                term *= (mu - (2.0 * k - 1.0) * (2.0 * k - 1.0)) / (k * 8.0 * System.Math.Abs(x));
                correction += term;

                if (System.Math.Abs(term) < epsilon)
                {
                    break;
                }
            }

            return result * correction;
        }
    }
}
