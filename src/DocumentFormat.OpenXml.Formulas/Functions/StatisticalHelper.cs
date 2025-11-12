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

    /// <summary>
    /// Calculates the natural logarithm of the gamma function using Lanczos approximation.
    /// </summary>
    public static double LogGamma(double x)
    {
        if (x <= 0)
        {
            throw new System.ArgumentException("Argument must be positive");
        }

        // Lanczos coefficients for g=7
        double[] coef = {
            0.99999999999980993,
            676.5203681218851,
            -1259.1392167224028,
            771.32342877765313,
            -176.61502916214059,
            12.507343278686905,
            -0.13857109526572012,
            9.9843695780195716e-6,
            1.5056327351493116e-7
        };

        double z = x;
        double sum = coef[0];
        for (int i = 1; i < 9; i++)
        {
            sum += coef[i] / (z + i);
        }

        double tmp = z + 7.5;
        return (z + 0.5) * System.Math.Log(tmp) - tmp + System.Math.Log(System.Math.Sqrt(2.0 * System.Math.PI) * sum / z);
    }

    /// <summary>
    /// Calculates the gamma function.
    /// </summary>
    public static double Gamma(double x)
    {
        return System.Math.Exp(LogGamma(x));
    }

    /// <summary>
    /// Calculates the incomplete beta function using continued fraction expansion.
    /// </summary>
    public static double IncompleteBeta(double x, double a, double b)
    {
        if (x < 0.0 || x > 1.0)
        {
            throw new System.ArgumentException("x must be between 0 and 1");
        }

        if (a <= 0.0 || b <= 0.0)
        {
            throw new System.ArgumentException("a and b must be positive");
        }

        if (x == 0.0) return 0.0;
        if (x == 1.0) return 1.0;

        // Use symmetry relation if necessary
        bool swap = false;
        if (x > (a + 1.0) / (a + b + 2.0))
        {
            swap = true;
            double temp = a;
            a = b;
            b = temp;
            x = 1.0 - x;
        }

        double logBeta = LogGamma(a) + LogGamma(b) - LogGamma(a + b);
        double front = System.Math.Exp(System.Math.Log(x) * a + System.Math.Log(1.0 - x) * b - logBeta) / a;

        // Continued fraction using modified Lentz's method
        double f = 1.0;
        double c = 1.0;
        double d = 0.0;

        for (int m = 0; m <= 200; m++)
        {
            double numerator, denominator;

            if (m == 0)
            {
                numerator = 1.0;
            }
            else if (m % 2 == 0)
            {
                int m2 = m / 2;
                numerator = (m2 * (b - m2) * x) / ((a + m - 1) * (a + m));
            }
            else
            {
                int m2 = (m - 1) / 2;
                numerator = -((a + m2) * (a + b + m2) * x) / ((a + m) * (a + m + 1));
            }

            denominator = 1.0;

            d = denominator + numerator * d;
            if (System.Math.Abs(d) < 1e-30) d = 1e-30;
            d = 1.0 / d;

            c = denominator + numerator / c;
            if (System.Math.Abs(c) < 1e-30) c = 1e-30;

            double delta = c * d;
            f *= delta;

            if (System.Math.Abs(delta - 1.0) < 3e-7)
                break;
        }

        double result = front * f;
        return swap ? 1.0 - result : result;
    }

    /// <summary>
    /// Calculates the regularized incomplete beta function I_x(a,b).
    /// </summary>
    public static double BetaCDF(double x, double a, double b)
    {
        return IncompleteBeta(x, a, b);
    }

    /// <summary>
    /// Calculates the incomplete gamma function (lower) using series expansion.
    /// </summary>
    public static double IncompleteGammaLower(double a, double x)
    {
        if (x < 0.0 || a <= 0.0)
        {
            throw new System.ArgumentException("Invalid arguments for incomplete gamma");
        }

        if (x == 0.0) return 0.0;

        // Use series expansion for lower incomplete gamma
        double sum = 1.0 / a;
        double term = 1.0 / a;

        for (int n = 1; n <= 200; n++)
        {
            term *= x / (a + n);
            sum += term;
            if (System.Math.Abs(term) < System.Math.Abs(sum) * 1e-10)
                break;
        }

        return sum * System.Math.Exp(-x + a * System.Math.Log(x) - LogGamma(a));
    }

    /// <summary>
    /// Calculates the incomplete gamma function (upper) using continued fraction.
    /// </summary>
    public static double IncompleteGammaUpper(double a, double x)
    {
        if (x < 0.0 || a <= 0.0)
        {
            throw new System.ArgumentException("Invalid arguments for incomplete gamma");
        }

        // Use continued fraction for upper incomplete gamma
        double b = x + 1.0 - a;
        double c = 1.0 / 1e-30;
        double d = 1.0 / b;
        double h = d;

        for (int i = 1; i <= 200; i++)
        {
            double an = -i * (i - a);
            b += 2.0;
            d = an * d + b;
            if (System.Math.Abs(d) < 1e-30) d = 1e-30;
            c = b + an / c;
            if (System.Math.Abs(c) < 1e-30) c = 1e-30;
            d = 1.0 / d;
            double delta = d * c;
            h *= delta;
            if (System.Math.Abs(delta - 1.0) < 1e-10)
                break;
        }

        return System.Math.Exp(-x + a * System.Math.Log(x) - LogGamma(a)) * h;
    }

    /// <summary>
    /// Calculates the regularized incomplete gamma function P(a,x).
    /// </summary>
    public static double GammaCDF(double x, double a)
    {
        if (x < 0.0)
        {
            return 0.0;
        }

        if (x < a + 1.0)
        {
            return IncompleteGammaLower(a, x) / Gamma(a);
        }
        else
        {
            return 1.0 - IncompleteGammaUpper(a, x) / Gamma(a);
        }
    }

    /// <summary>
    /// Calculates Student's t-distribution CDF using the relationship with incomplete beta function.
    /// </summary>
    public static double TDistCDF(double t, double df)
    {
        if (df <= 0)
        {
            throw new System.ArgumentException("Degrees of freedom must be positive");
        }

        double x = df / (df + t * t);
        double beta = 0.5 * IncompleteBeta(x, df / 2.0, 0.5);

        if (t > 0)
        {
            return 1.0 - beta;
        }
        else
        {
            return beta;
        }
    }

    /// <summary>
    /// Calculates Student's t-distribution PDF.
    /// </summary>
    public static double TDistPDF(double t, double df)
    {
        if (df <= 0)
        {
            throw new System.ArgumentException("Degrees of freedom must be positive");
        }

        double numerator = Gamma((df + 1.0) / 2.0);
        double denominator = System.Math.Sqrt(df * System.Math.PI) * Gamma(df / 2.0);
        double factor = System.Math.Pow(1.0 + (t * t) / df, -(df + 1.0) / 2.0);

        return (numerator / denominator) * factor;
    }

    /// <summary>
    /// Calculates the inverse of Student's t-distribution using Newton-Raphson method.
    /// </summary>
    public static double TDistInv(double p, double df)
    {
        if (p <= 0.0 || p >= 1.0)
        {
            throw new System.ArgumentException("Probability must be between 0 and 1 (exclusive)");
        }

        if (df <= 0)
        {
            throw new System.ArgumentException("Degrees of freedom must be positive");
        }

        // For large df, t-distribution approaches normal distribution
        if (df > 1000)
        {
            return NormSInv(p);
        }

        // Initial guess using normal approximation
        double t = NormSInv(p);

        // Newton-Raphson iteration
        for (int i = 0; i < 10; i++)
        {
            double cdf = TDistCDF(t, df);
            double pdf = TDistPDF(t, df);

            if (System.Math.Abs(pdf) < 1e-20)
                break;

            double delta = (cdf - p) / pdf;
            t -= delta;

            if (System.Math.Abs(delta) < 1e-8)
                break;
        }

        return t;
    }

    /// <summary>
    /// Calculates the chi-squared distribution CDF.
    /// </summary>
    public static double ChiSquareCDF(double x, double df)
    {
        if (x < 0.0)
        {
            return 0.0;
        }

        if (df <= 0)
        {
            throw new System.ArgumentException("Degrees of freedom must be positive");
        }

        return GammaCDF(x / 2.0, df / 2.0);
    }

    /// <summary>
    /// Calculates the chi-squared distribution PDF.
    /// </summary>
    public static double ChiSquarePDF(double x, double df)
    {
        if (x < 0.0)
        {
            return 0.0;
        }

        if (df <= 0)
        {
            throw new System.ArgumentException("Degrees of freedom must be positive");
        }

        if (x == 0.0)
        {
            if (df < 2.0)
                return double.PositiveInfinity;
            else if (df == 2.0)
                return 0.5;
            else
                return 0.0;
        }

        double k = df / 2.0;
        return System.Math.Exp((k - 1.0) * System.Math.Log(x) - x / 2.0 - k * System.Math.Log(2.0) - LogGamma(k));
    }

    /// <summary>
    /// Calculates the inverse of the chi-squared distribution using Newton-Raphson method.
    /// </summary>
    public static double ChiSquareInv(double p, double df)
    {
        if (p <= 0.0 || p >= 1.0)
        {
            throw new System.ArgumentException("Probability must be between 0 and 1 (exclusive)");
        }

        if (df <= 0)
        {
            throw new System.ArgumentException("Degrees of freedom must be positive");
        }

        // Initial guess using Wilson-Hilferty transformation
        double z = NormSInv(p);
        double x = df * System.Math.Pow(1.0 - 2.0 / (9.0 * df) + z * System.Math.Sqrt(2.0 / (9.0 * df)), 3.0);

        if (x < 0) x = 0.001;

        // Newton-Raphson iteration
        for (int i = 0; i < 10; i++)
        {
            double cdf = ChiSquareCDF(x, df);
            double pdf = ChiSquarePDF(x, df);

            if (System.Math.Abs(pdf) < 1e-20)
                break;

            double delta = (cdf - p) / pdf;
            x -= delta;

            if (x < 0) x = 0.001;

            if (System.Math.Abs(delta) < 1e-8)
                break;
        }

        return x;
    }

    /// <summary>
    /// Calculates the F-distribution CDF using the relationship with incomplete beta function.
    /// </summary>
    public static double FDistCDF(double x, double df1, double df2)
    {
        if (x < 0.0)
        {
            return 0.0;
        }

        if (df1 <= 0 || df2 <= 0)
        {
            throw new System.ArgumentException("Degrees of freedom must be positive");
        }

        double t = df2 / (df2 + df1 * x);
        return 1.0 - IncompleteBeta(t, df2 / 2.0, df1 / 2.0);
    }

    /// <summary>
    /// Calculates the F-distribution PDF.
    /// </summary>
    public static double FDistPDF(double x, double df1, double df2)
    {
        if (x < 0.0)
        {
            return 0.0;
        }

        if (df1 <= 0 || df2 <= 0)
        {
            throw new System.ArgumentException("Degrees of freedom must be positive");
        }

        if (x == 0.0)
        {
            if (df1 < 2.0)
                return double.PositiveInfinity;
            else if (df1 == 2.0)
                return 1.0;
            else
                return 0.0;
        }

        double a = df1 / 2.0;
        double b = df2 / 2.0;
        double numerator = System.Math.Pow(df1 * x, a) * System.Math.Pow(df2, b);
        double denominator = System.Math.Pow(df2 + df1 * x, a + b) * x;

        return (numerator / denominator) * System.Math.Exp(LogGamma(a + b) - LogGamma(a) - LogGamma(b));
    }

    /// <summary>
    /// Calculates the inverse of the F-distribution using Newton-Raphson method.
    /// </summary>
    public static double FDistInv(double p, double df1, double df2)
    {
        if (p <= 0.0 || p >= 1.0)
        {
            throw new System.ArgumentException("Probability must be between 0 and 1 (exclusive)");
        }

        if (df1 <= 0 || df2 <= 0)
        {
            throw new System.ArgumentException("Degrees of freedom must be positive");
        }

        // Initial guess
        double x = 1.0;

        // Newton-Raphson iteration
        for (int i = 0; i < 20; i++)
        {
            double cdf = FDistCDF(x, df1, df2);
            double pdf = FDistPDF(x, df1, df2);

            if (System.Math.Abs(pdf) < 1e-20)
                break;

            double delta = (cdf - p) / pdf;
            x -= delta;

            if (x < 0.0001) x = 0.0001;

            if (System.Math.Abs(delta) < 1e-8)
                break;
        }

        return x;
    }

    /// <summary>
    /// Calculates the beta distribution PDF.
    /// </summary>
    public static double BetaPDF(double x, double alpha, double beta)
    {
        if (x < 0.0 || x > 1.0)
        {
            return 0.0;
        }

        if (alpha <= 0 || beta <= 0)
        {
            throw new System.ArgumentException("Alpha and beta must be positive");
        }

        if (x == 0.0)
        {
            if (alpha < 1.0) return double.PositiveInfinity;
            else if (alpha == 1.0) return beta;
            else return 0.0;
        }

        if (x == 1.0)
        {
            if (beta < 1.0) return double.PositiveInfinity;
            else if (beta == 1.0) return alpha;
            else return 0.0;
        }

        return System.Math.Exp((alpha - 1.0) * System.Math.Log(x) + (beta - 1.0) * System.Math.Log(1.0 - x) +
                              LogGamma(alpha + beta) - LogGamma(alpha) - LogGamma(beta));
    }

    /// <summary>
    /// Calculates the inverse of the beta distribution using Newton-Raphson method.
    /// </summary>
    public static double BetaInv(double p, double alpha, double beta)
    {
        if (p < 0.0 || p > 1.0)
        {
            throw new System.ArgumentException("Probability must be between 0 and 1");
        }

        if (alpha <= 0 || beta <= 0)
        {
            throw new System.ArgumentException("Alpha and beta must be positive");
        }

        if (p == 0.0) return 0.0;
        if (p == 1.0) return 1.0;

        // Initial guess
        double x = alpha / (alpha + beta);

        // Newton-Raphson iteration
        for (int i = 0; i < 20; i++)
        {
            double cdf = BetaCDF(x, alpha, beta);
            double pdf = BetaPDF(x, alpha, beta);

            if (System.Math.Abs(pdf) < 1e-20)
                break;

            double delta = (cdf - p) / pdf;
            x -= delta;

            if (x < 0.0) x = 0.0001;
            if (x > 1.0) x = 0.9999;

            if (System.Math.Abs(delta) < 1e-8)
                break;
        }

        return x;
    }

    /// <summary>
    /// Calculates the lognormal distribution CDF.
    /// </summary>
    public static double LogNormCDF(double x, double mean, double stdDev)
    {
        if (x <= 0.0)
        {
            return 0.0;
        }

        if (stdDev <= 0)
        {
            throw new System.ArgumentException("Standard deviation must be positive");
        }

        double z = (System.Math.Log(x) - mean) / stdDev;
        return NormSDist(z);
    }

    /// <summary>
    /// Calculates the lognormal distribution PDF.
    /// </summary>
    public static double LogNormPDF(double x, double mean, double stdDev)
    {
        if (x <= 0.0)
        {
            return 0.0;
        }

        if (stdDev <= 0)
        {
            throw new System.ArgumentException("Standard deviation must be positive");
        }

        double z = (System.Math.Log(x) - mean) / stdDev;
        return System.Math.Exp(-0.5 * z * z) / (x * stdDev * System.Math.Sqrt(2.0 * System.Math.PI));
    }

    /// <summary>
    /// Calculates the inverse of the lognormal distribution.
    /// </summary>
    public static double LogNormInv(double p, double mean, double stdDev)
    {
        if (p <= 0.0 || p >= 1.0)
        {
            throw new System.ArgumentException("Probability must be between 0 and 1 (exclusive)");
        }

        if (stdDev <= 0)
        {
            throw new System.ArgumentException("Standard deviation must be positive");
        }

        double z = NormSInv(p);
        return System.Math.Exp(mean + stdDev * z);
    }

    /// <summary>
    /// Calculates the binomial coefficient C(n,k) = n! / (k! * (n-k)!).
    /// </summary>
    public static double BinomialCoefficient(int n, int k)
    {
        if (k < 0 || k > n)
        {
            return 0;
        }

        if (k == 0 || k == n)
        {
            return 1;
        }

        // Use symmetry property
        if (k > n - k)
        {
            k = n - k;
        }

        // Calculate using logarithms to avoid overflow
        double result = 0;
        for (int i = 0; i < k; i++)
        {
            result += System.Math.Log(n - i) - System.Math.Log(i + 1);
        }

        return System.Math.Exp(result);
    }

    /// <summary>
    /// Calculates the binomial probability mass function.
    /// </summary>
    public static double BinomialPMF(int k, int n, double p)
    {
        if (k < 0 || k > n || p < 0 || p > 1)
        {
            return 0;
        }

        if (p == 0)
        {
            return k == 0 ? 1.0 : 0.0;
        }

        if (p == 1)
        {
            return k == n ? 1.0 : 0.0;
        }

        // Use logarithms for numerical stability
        double logProb = System.Math.Log(BinomialCoefficient(n, k)) +
                         k * System.Math.Log(p) +
                         (n - k) * System.Math.Log(1 - p);

        return System.Math.Exp(logProb);
    }

    /// <summary>
    /// Calculates the binomial cumulative distribution function.
    /// </summary>
    public static double BinomialCDF(int k, int n, double p)
    {
        if (k < 0)
        {
            return 0;
        }

        if (k >= n)
        {
            return 1;
        }

        if (p == 0)
        {
            return 1.0;
        }

        if (p == 1)
        {
            return k >= n ? 1.0 : 0.0;
        }

        // Sum probabilities from 0 to k
        double sum = 0;
        for (int i = 0; i <= k; i++)
        {
            sum += BinomialPMF(i, n, p);
        }

        return sum;
    }

    /// <summary>
    /// Calculates the Poisson probability mass function.
    /// </summary>
    public static double PoissonPMF(int k, double lambda)
    {
        if (k < 0 || lambda <= 0)
        {
            return 0;
        }

        // Use logarithms for numerical stability
        double logProb = k * System.Math.Log(lambda) - lambda - LogGamma(k + 1);
        return System.Math.Exp(logProb);
    }

    /// <summary>
    /// Calculates the Poisson cumulative distribution function.
    /// </summary>
    public static double PoissonCDF(int k, double lambda)
    {
        if (k < 0)
        {
            return 0;
        }

        if (lambda <= 0)
        {
            throw new System.ArgumentException("Lambda must be positive");
        }

        // Sum probabilities from 0 to k
        double sum = 0;
        for (int i = 0; i <= k; i++)
        {
            sum += PoissonPMF(i, lambda);
        }

        return sum;
    }

    /// <summary>
    /// Calculates the exponential probability density function.
    /// </summary>
    public static double ExponentialPDF(double x, double lambda)
    {
        if (x < 0 || lambda <= 0)
        {
            return 0;
        }

        return lambda * System.Math.Exp(-lambda * x);
    }

    /// <summary>
    /// Calculates the exponential cumulative distribution function.
    /// </summary>
    public static double ExponentialCDF(double x, double lambda)
    {
        if (x < 0)
        {
            return 0;
        }

        if (lambda <= 0)
        {
            throw new System.ArgumentException("Lambda must be positive");
        }

        return 1.0 - System.Math.Exp(-lambda * x);
    }

    /// <summary>
    /// Calculates the Weibull probability density function.
    /// </summary>
    public static double WeibullPDF(double x, double alpha, double beta)
    {
        if (x < 0 || alpha <= 0 || beta <= 0)
        {
            return 0;
        }

        double xOverBeta = x / beta;
        return (alpha / beta) * System.Math.Pow(xOverBeta, alpha - 1) *
               System.Math.Exp(-System.Math.Pow(xOverBeta, alpha));
    }

    /// <summary>
    /// Calculates the Weibull cumulative distribution function.
    /// </summary>
    public static double WeibullCDF(double x, double alpha, double beta)
    {
        if (x < 0)
        {
            return 0;
        }

        if (alpha <= 0 || beta <= 0)
        {
            throw new System.ArgumentException("Alpha and beta must be positive");
        }

        return 1.0 - System.Math.Exp(-System.Math.Pow(x / beta, alpha));
    }

    /// <summary>
    /// Calculates the gamma probability density function.
    /// </summary>
    public static double GammaPDF(double x, double alpha, double beta)
    {
        if (x < 0 || alpha <= 0 || beta <= 0)
        {
            return 0;
        }

        if (x == 0)
        {
            if (alpha < 1)
                return double.PositiveInfinity;
            else if (alpha == 1)
                return 1.0 / beta;
            else
                return 0;
        }

        double logProb = (alpha - 1) * System.Math.Log(x) - x / beta -
                         alpha * System.Math.Log(beta) - LogGamma(alpha);

        return System.Math.Exp(logProb);
    }

    /// <summary>
    /// Calculates the gamma cumulative distribution function.
    /// </summary>
    public static double GammaDistCDF(double x, double alpha, double beta)
    {
        if (x < 0)
        {
            return 0;
        }

        if (alpha <= 0 || beta <= 0)
        {
            throw new System.ArgumentException("Alpha and beta must be positive");
        }

        return GammaCDF(x / beta, alpha);
    }

    /// <summary>
    /// Calculates the inverse of the gamma distribution using Newton-Raphson method.
    /// </summary>
    public static double GammaInv(double p, double alpha, double beta)
    {
        if (p < 0.0 || p > 1.0)
        {
            throw new System.ArgumentException("Probability must be between 0 and 1");
        }

        if (alpha <= 0 || beta <= 0)
        {
            throw new System.ArgumentException("Alpha and beta must be positive");
        }

        if (p == 0.0) return 0.0;
        if (p == 1.0) return double.PositiveInfinity;

        // Initial guess
        double x = alpha * beta;

        // Newton-Raphson iteration
        for (int i = 0; i < 20; i++)
        {
            double cdf = GammaDistCDF(x, alpha, beta);
            double pdf = GammaPDF(x, alpha, beta);

            if (System.Math.Abs(pdf) < 1e-20)
                break;

            double delta = (cdf - p) / pdf;
            x -= delta;

            if (x < 0) x = 0.0001;

            if (System.Math.Abs(delta) < 1e-8)
                break;
        }

        return x;
    }
}
