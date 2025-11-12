// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the BESSELK function.
/// BESSELK(x, n) - returns the modified Bessel function Kn(x).
/// </summary>
public sealed class BesselkFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly BesselkFunction Instance = new();

    private BesselkFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "BESSELK";

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

        if (n < 0 || x <= 0)
        {
            return CellValue.Error("#NUM!");
        }

        try
        {
            var result = BesselK(x, n);
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
    /// Computes the modified Bessel function of the second kind Kn(x).
    /// Uses recurrence relation.
    /// </summary>
    private static double BesselK(double x, int n)
    {
        if (n == 0)
        {
            return BesselK0(x);
        }

        if (n == 1)
        {
            return BesselK1(x);
        }

        // Use recurrence relation for n >= 2
        // K_{n+1}(x) = K_{n-1}(x) + (2n/x) * K_n(x)
        double k0 = BesselK0(x);
        double k1 = BesselK1(x);

        for (int i = 1; i < n; i++)
        {
            double kNext = k0 + (2.0 * i / x) * k1;
            k0 = k1;
            k1 = kNext;
        }

        return k1;
    }

    private static double BesselK0(double x)
    {
        if (x <= 2.0)
        {
            double y = x * x / 4.0;
            double ans = (-System.Math.Log(x / 2.0) * BesselI0(x)) + (-0.57721566 + y * (0.42278420
                + y * (0.23069756 + y * (0.3488590e-1 + y * (0.262698e-2
                + y * (0.10750e-3 + y * 0.74e-5))))));
            return ans;
        }
        else
        {
            double y = 2.0 / x;
            double ans = (System.Math.Exp(-x) / System.Math.Sqrt(x)) * (1.25331414 + y * (-0.7832358e-1
                + y * (0.2189568e-1 + y * (-0.1062446e-1 + y * (0.587872e-2
                + y * (-0.251540e-2 + y * 0.53208e-3))))));
            return ans;
        }
    }

    private static double BesselK1(double x)
    {
        if (x <= 2.0)
        {
            double y = x * x / 4.0;
            double ans = (System.Math.Log(x / 2.0) * BesselI1(x)) + (1.0 / x) * (1.0 + y * (0.15443144
                + y * (-0.67278579 + y * (-0.18156897 + y * (-0.1919402e-1
                + y * (-0.110404e-2 + y * (-0.4686e-4)))))));
            return ans;
        }
        else
        {
            double y = 2.0 / x;
            double ans = (System.Math.Exp(-x) / System.Math.Sqrt(x)) * (1.25331414 + y * (0.23498619
                + y * (-0.3655620e-1 + y * (0.1504268e-1 + y * (-0.780353e-2
                + y * (0.325614e-2 + y * (-0.68245e-3)))))));
            return ans;
        }
    }

    private static double BesselI0(double x)
    {
        double ax = System.Math.Abs(x);
        if (ax < 3.75)
        {
            double y = x / 3.75;
            y *= y;
            return 1.0 + y * (3.5156229 + y * (3.0899424 + y * (1.2067492
                + y * (0.2659732 + y * (0.360768e-1 + y * 0.45813e-2)))));
        }
        else
        {
            double y = 3.75 / ax;
            return (System.Math.Exp(ax) / System.Math.Sqrt(ax)) * (0.39894228 + y * (0.1328592e-1
                + y * (0.225319e-2 + y * (-0.157565e-2 + y * (0.916281e-2
                + y * (-0.2057706e-1 + y * (0.2635537e-1 + y * (-0.1647633e-1
                + y * 0.392377e-2))))))));
        }
    }

    private static double BesselI1(double x)
    {
        double ax = System.Math.Abs(x);
        if (ax < 3.75)
        {
            double y = x / 3.75;
            y *= y;
            return ax * (0.5 + y * (0.87890594 + y * (0.51498869 + y * (0.15084934
                + y * (0.2658733e-1 + y * (0.301532e-2 + y * 0.32411e-3))))));
        }
        else
        {
            double y = 3.75 / ax;
            double ans = 0.2282967e-1 + y * (-0.2895312e-1 + y * (0.1787654e-1
                - y * 0.420059e-2));
            ans = 0.39894228 + y * (-0.3988024e-1 + y * (-0.362018e-2
                + y * (0.163801e-2 + y * (-0.1031555e-1 + y * ans))));
            ans *= (System.Math.Exp(ax) / System.Math.Sqrt(ax));
            return x < 0.0 ? -ans : ans;
        }
    }
}
