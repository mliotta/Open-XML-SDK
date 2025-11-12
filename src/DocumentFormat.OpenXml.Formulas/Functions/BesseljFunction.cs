// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the BESSELJ function.
/// BESSELJ(x, n) - returns the Bessel function Jn(x).
/// </summary>
public sealed class BesseljFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly BesseljFunction Instance = new();

    private BesseljFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "BESSELJ";

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
            var result = BesselJ(x, n);
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
    /// Computes the Bessel function of the first kind Jn(x).
    /// Uses series expansion for small x and recurrence relation for higher orders.
    /// </summary>
    private static double BesselJ(double x, int n)
    {
        const int maxIterations = 100;
        const double epsilon = 1e-10;

        if (n == 0)
        {
            return BesselJ0(x);
        }

        if (n == 1)
        {
            return BesselJ1(x);
        }

        // Use recurrence relation for n >= 2
        // J_{n+1}(x) = (2n/x) * J_n(x) - J_{n-1}(x)
        double j0 = BesselJ0(x);
        double j1 = BesselJ1(x);

        if (System.Math.Abs(x) < epsilon)
        {
            return 0.0;
        }

        for (int i = 1; i < n; i++)
        {
            double jNext = (2.0 * i / x) * j1 - j0;
            j0 = j1;
            j1 = jNext;
        }

        return j1;
    }

    private static double BesselJ0(double x)
    {
        const double epsilon = 1e-10;
        double ax = System.Math.Abs(x);

        if (ax < 8.0)
        {
            // Series expansion for small x
            double y = x * x;
            double ans1 = 57568490574.0 + y * (-13362590354.0 + y * (651619640.7
                + y * (-11214424.18 + y * (77392.33017 + y * (-184.9052456)))));
            double ans2 = 57568490411.0 + y * (1029532985.0 + y * (9494680.718
                + y * (59272.64853 + y * (267.8532712 + y * 1.0))));
            return ans1 / ans2;
        }
        else
        {
            // Asymptotic expansion for large x
            double z = 8.0 / ax;
            double y = z * z;
            double xx = ax - 0.785398164;
            double ans1 = 1.0 + y * (-0.1098628627e-2 + y * (0.2734510407e-4
                + y * (-0.2073370639e-5 + y * 0.2093887211e-6)));
            double ans2 = -0.1562499995e-1 + y * (0.1430488765e-3
                + y * (-0.6911147651e-5 + y * (0.7621095161e-6
                - y * 0.934935152e-7)));
            return System.Math.Sqrt(0.636619772 / ax) *
                (System.Math.Cos(xx) * ans1 - z * System.Math.Sin(xx) * ans2);
        }
    }

    private static double BesselJ1(double x)
    {
        double ax = System.Math.Abs(x);

        if (ax < 8.0)
        {
            // Series expansion for small x
            double y = x * x;
            double ans1 = x * (72362614232.0 + y * (-7895059235.0 + y * (242396853.1
                + y * (-2972611.439 + y * (15704.48260 + y * (-30.16036606))))));
            double ans2 = 144725228442.0 + y * (2300535178.0 + y * (18583304.74
                + y * (99447.43394 + y * (376.9991397 + y * 1.0))));
            return ans1 / ans2;
        }
        else
        {
            // Asymptotic expansion for large x
            double z = 8.0 / ax;
            double y = z * z;
            double xx = ax - 2.356194491;
            double ans1 = 1.0 + y * (0.183105e-2 + y * (-0.3516396496e-4
                + y * (0.2457520174e-5 + y * (-0.240337019e-6))));
            double ans2 = 0.04687499995 + y * (-0.2002690873e-3
                + y * (0.8449199096e-5 + y * (-0.88228987e-6
                + y * 0.105787412e-6)));
            double ans = System.Math.Sqrt(0.636619772 / ax) *
                (System.Math.Cos(xx) * ans1 - z * System.Math.Sin(xx) * ans2);
            return x < 0.0 ? -ans : ans;
        }
    }
}
