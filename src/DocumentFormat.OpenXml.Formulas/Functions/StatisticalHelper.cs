// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Helper class for statistical distribution calculations.
/// </summary>
internal static class StatisticalHelper
{
    /// <summary>
    /// Calculates the standard normal cumulative distribution function (CDF).
    /// Uses the error function approximation.
    /// </summary>
    public static double NormSDist(double z)
    {
        return 0.5 * (1.0 + Erf(z / System.Math.Sqrt(2.0)));
    }

    /// <summary>
    /// Calculates the standard normal probability density function (PDF).
    /// </summary>
    public static double NormSPdf(double z)
    {
        return System.Math.Exp(-0.5 * z * z) / System.Math.Sqrt(2.0 * System.Math.PI);
    }

    /// <summary>
    /// Calculates the inverse of the standard normal cumulative distribution function.
    /// Uses Beasley-Springer-Moro algorithm.
    /// </summary>
    public static double NormSInv(double p)
    {
        if (p <= 0.0 || p >= 1.0)
        {
            throw new System.ArgumentException("Probability must be between 0 and 1 (exclusive)");
        }

        // Coefficients for the rational approximation
        double[] a = { -3.969683028665376e+01, 2.209460984245205e+02,
                       -2.759285104469687e+02, 1.383577518672690e+02,
                       -3.066479806614716e+01, 2.506628277459239e+00 };

        double[] b = { -5.447609879822406e+01, 1.615858368580409e+02,
                       -1.556989798598866e+02, 6.680131188771972e+01,
                       -1.328068155288572e+01 };

        double[] c = { -7.784894002430293e-03, -3.223964580411365e-01,
                       -2.400758277161838e+00, -2.549732539343734e+00,
                        4.374664141464968e+00,  2.938163982698783e+00 };

        double[] d = { 7.784695709041462e-03, 3.224671290700398e-01,
                       2.445134137142996e+00, 3.754408661907416e+00 };

        double q, r, result;

        if (p < 0.02425)
        {
            q = System.Math.Sqrt(-2.0 * System.Math.Log(p));
            result = (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) /
                     ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1.0);
        }
        else if (p > 0.97575)
        {
            q = System.Math.Sqrt(-2.0 * System.Math.Log(1.0 - p));
            result = -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) /
                      ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1.0);
        }
        else
        {
            q = p - 0.5;
            r = q * q;
            result = (((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q /
                     (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1.0);
        }

        return result;
    }

    /// <summary>
    /// Error function approximation using Abramowitz and Stegun formula.
    /// </summary>
    private static double Erf(double x)
    {
        // Constants
        const double a1 = 0.254829592;
        const double a2 = -0.284496736;
        const double a3 = 1.421413741;
        const double a4 = -1.453152027;
        const double a5 = 1.061405429;
        const double p = 0.3275911;

        // Save the sign of x
        int sign = x < 0 ? -1 : 1;
        x = System.Math.Abs(x);

        // A&S formula 7.1.26
        double t = 1.0 / (1.0 + p * x);
        double y = 1.0 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * System.Math.Exp(-x * x);

        return sign * y;
    }

    /// <summary>
    /// Calculates the normal cumulative distribution function with mean and standard deviation.
    /// </summary>
    public static double NormDist(double x, double mean, double standardDev, bool cumulative)
    {
        if (standardDev <= 0)
        {
            throw new System.ArgumentException("Standard deviation must be positive");
        }

        double z = (x - mean) / standardDev;

        if (cumulative)
        {
            return NormSDist(z);
        }
        else
        {
            // Probability density function
            return NormSPdf(z) / standardDev;
        }
    }

    /// <summary>
    /// Calculates the inverse of the normal cumulative distribution function.
    /// </summary>
    public static double NormInv(double probability, double mean, double standardDev)
    {
        if (probability <= 0.0 || probability >= 1.0)
        {
            throw new System.ArgumentException("Probability must be between 0 and 1 (exclusive)");
        }

        if (standardDev <= 0)
        {
            throw new System.ArgumentException("Standard deviation must be positive");
        }

        double z = NormSInv(probability);
        return mean + z * standardDev;
    }
}
