// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Helper class for complex number operations.
/// Handles complex numbers in the format "a+bi" or "a+bj".
/// </summary>
internal sealed class ComplexNumber
{
    public double Real { get; set; }

    public double Imaginary { get; set; }

    public ComplexNumber(double real, double imaginary)
    {
        Real = real;
        Imaginary = imaginary;
    }

    /// <summary>
    /// Parses a complex number string in the format "a+bi" or "a+bj".
    /// </summary>
    public static bool TryParse(string value, out ComplexNumber? result)
    {
        result = null;

        if (string.IsNullOrEmpty(value) || value.Trim().Length == 0)
        {
            return false;
        }

        value = value.Trim().Replace(" ", string.Empty);

        // Check for suffix (i or j)
        var suffix = value.EndsWith("i") ? "i" : value.EndsWith("j") ? "j" : null;
        if (suffix == null)
        {
            return false;
        }

        value = value.Substring(0, value.Length - 1);

        // Pure imaginary number (e.g., "5i", "-3j")
        if (!value.Contains("+") && !(value.IndexOf("-", 1, StringComparison.Ordinal) >= 0))
        {
            if (string.IsNullOrEmpty(value) || value == "+" || value == "-")
            {
                // Just "i" or "j" means 1i
                result = new ComplexNumber(0, value == "-" ? -1 : 1);
                return true;
            }

            if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var img))
            {
                result = new ComplexNumber(0, img);
                return true;
            }

            return false;
        }

        // Find the position of the last + or - (which separates real and imaginary parts)
        int splitPos = -1;
        for (int i = value.Length - 1; i > 0; i--)
        {
            if (value[i] == '+' || value[i] == '-')
            {
                splitPos = i;
                break;
            }
        }

        if (splitPos <= 0)
        {
            return false;
        }

        var realPart = value.Substring(0, splitPos);
        var imagPart = value.Substring(splitPos);

        if (string.IsNullOrEmpty(realPart))
        {
            return false;
        }

        if (!double.TryParse(realPart, NumberStyles.Float, CultureInfo.InvariantCulture, out var real))
        {
            return false;
        }

        // Handle cases where imaginary part is just "+" or "-" or "+i" or "-i"
        if (imagPart == "+" || imagPart == "+1" || string.IsNullOrEmpty(imagPart.Substring(1)))
        {
            result = new ComplexNumber(real, 1);
            return true;
        }

        if (imagPart == "-" || imagPart == "-1")
        {
            result = new ComplexNumber(real, -1);
            return true;
        }

        if (!double.TryParse(imagPart, NumberStyles.Float, CultureInfo.InvariantCulture, out var imag))
        {
            return false;
        }

        result = new ComplexNumber(real, imag);
        return true;
    }

    /// <summary>
    /// Converts the complex number to string format.
    /// </summary>
    public string ToString(string suffix)
    {
        if (System.Math.Abs(Imaginary) < 1e-10)
        {
            return Real.ToString(CultureInfo.InvariantCulture);
        }

        if (System.Math.Abs(Real) < 1e-10)
        {
            if (System.Math.Abs(Imaginary - 1.0) < 1e-10)
            {
                return suffix;
            }

            if (System.Math.Abs(Imaginary + 1.0) < 1e-10)
            {
                return "-" + suffix;
            }

            return Imaginary.ToString(CultureInfo.InvariantCulture) + suffix;
        }

        var imagStr = string.Empty;
        if (System.Math.Abs(Imaginary - 1.0) < 1e-10)
        {
            imagStr = "+" + suffix;
        }
        else if (System.Math.Abs(Imaginary + 1.0) < 1e-10)
        {
            imagStr = "-" + suffix;
        }
        else if (Imaginary > 0)
        {
            imagStr = "+" + Imaginary.ToString(CultureInfo.InvariantCulture) + suffix;
        }
        else
        {
            imagStr = Imaginary.ToString(CultureInfo.InvariantCulture) + suffix;
        }

        return Real.ToString(CultureInfo.InvariantCulture) + imagStr;
    }

    /// <summary>
    /// Returns the absolute value (modulus) of the complex number.
    /// </summary>
    public double Abs()
    {
        return System.Math.Sqrt(Real * Real + Imaginary * Imaginary);
    }

    /// <summary>
    /// Returns the argument (angle) of the complex number in radians.
    /// </summary>
    public double Argument()
    {
        return System.Math.Atan2(Imaginary, Real);
    }

    /// <summary>
    /// Returns the complex conjugate.
    /// </summary>
    public ComplexNumber Conjugate()
    {
        return new ComplexNumber(Real, -Imaginary);
    }

    /// <summary>
    /// Adds two complex numbers.
    /// </summary>
    public static ComplexNumber Add(ComplexNumber a, ComplexNumber b)
    {
        return new ComplexNumber(a.Real + b.Real, a.Imaginary + b.Imaginary);
    }

    /// <summary>
    /// Subtracts two complex numbers.
    /// </summary>
    public static ComplexNumber Subtract(ComplexNumber a, ComplexNumber b)
    {
        return new ComplexNumber(a.Real - b.Real, a.Imaginary - b.Imaginary);
    }

    /// <summary>
    /// Multiplies two complex numbers.
    /// </summary>
    public static ComplexNumber Multiply(ComplexNumber a, ComplexNumber b)
    {
        var real = a.Real * b.Real - a.Imaginary * b.Imaginary;
        var imag = a.Real * b.Imaginary + a.Imaginary * b.Real;
        return new ComplexNumber(real, imag);
    }

    /// <summary>
    /// Divides two complex numbers.
    /// </summary>
    public static ComplexNumber Divide(ComplexNumber a, ComplexNumber b)
    {
        var denominator = b.Real * b.Real + b.Imaginary * b.Imaginary;
        if (System.Math.Abs(denominator) < 1e-10)
        {
            return new ComplexNumber(double.NaN, double.NaN);
        }

        var real = (a.Real * b.Real + a.Imaginary * b.Imaginary) / denominator;
        var imag = (a.Imaginary * b.Real - a.Real * b.Imaginary) / denominator;
        return new ComplexNumber(real, imag);
    }

    /// <summary>
    /// Raises a complex number to an integer power.
    /// </summary>
    public static ComplexNumber Power(ComplexNumber z, int n)
    {
        if (n == 0)
        {
            return new ComplexNumber(1, 0);
        }

        if (n < 0)
        {
            return Divide(new ComplexNumber(1, 0), Power(z, -n));
        }

        var result = new ComplexNumber(1, 0);
        var current = z;

        while (n > 0)
        {
            if ((n & 1) == 1)
            {
                result = Multiply(result, current);
            }

            current = Multiply(current, current);
            n >>= 1;
        }

        return result;
    }

    /// <summary>
    /// Returns the square root of a complex number.
    /// </summary>
    public static ComplexNumber Sqrt(ComplexNumber z)
    {
        var r = z.Abs();
        var real = System.Math.Sqrt((r + z.Real) / 2.0);
        var imag = System.Math.Sign(z.Imaginary) * System.Math.Sqrt((r - z.Real) / 2.0);
        return new ComplexNumber(real, imag);
    }

    /// <summary>
    /// Returns the exponential of a complex number.
    /// </summary>
    public static ComplexNumber Exp(ComplexNumber z)
    {
        var expReal = System.Math.Exp(z.Real);
        return new ComplexNumber(expReal * System.Math.Cos(z.Imaginary), expReal * System.Math.Sin(z.Imaginary));
    }

    /// <summary>
    /// Returns the natural logarithm of a complex number.
    /// </summary>
    public static ComplexNumber Ln(ComplexNumber z)
    {
        return new ComplexNumber(System.Math.Log(z.Abs()), z.Argument());
    }

    /// <summary>
    /// Returns the base-10 logarithm of a complex number.
    /// </summary>
    public static ComplexNumber Log10(ComplexNumber z)
    {
        var ln = Ln(z);
        var log10e = System.Math.Log10(System.Math.E);
        return new ComplexNumber(ln.Real * log10e, ln.Imaginary * log10e);
    }

    /// <summary>
    /// Returns the base-2 logarithm of a complex number.
    /// </summary>
    public static ComplexNumber Log2(ComplexNumber z)
    {
        var ln = Ln(z);
        var log2e = System.Math.Log(2);
        return new ComplexNumber(ln.Real / log2e, ln.Imaginary / log2e);
    }

    /// <summary>
    /// Returns the sine of a complex number.
    /// </summary>
    public static ComplexNumber Sin(ComplexNumber z)
    {
        var real = System.Math.Sin(z.Real) * System.Math.Cosh(z.Imaginary);
        var imag = System.Math.Cos(z.Real) * System.Math.Sinh(z.Imaginary);
        return new ComplexNumber(real, imag);
    }

    /// <summary>
    /// Returns the cosine of a complex number.
    /// </summary>
    public static ComplexNumber Cos(ComplexNumber z)
    {
        var real = System.Math.Cos(z.Real) * System.Math.Cosh(z.Imaginary);
        var imag = -System.Math.Sin(z.Real) * System.Math.Sinh(z.Imaginary);
        return new ComplexNumber(real, imag);
    }

    /// <summary>
    /// Returns the tangent of a complex number.
    /// </summary>
    public static ComplexNumber Tan(ComplexNumber z)
    {
        return Divide(Sin(z), Cos(z));
    }

    /// <summary>
    /// Returns the secant of a complex number.
    /// </summary>
    public static ComplexNumber Sec(ComplexNumber z)
    {
        return Divide(new ComplexNumber(1, 0), Cos(z));
    }

    /// <summary>
    /// Returns the cosecant of a complex number.
    /// </summary>
    public static ComplexNumber Csc(ComplexNumber z)
    {
        return Divide(new ComplexNumber(1, 0), Sin(z));
    }

    /// <summary>
    /// Returns the cotangent of a complex number.
    /// </summary>
    public static ComplexNumber Cot(ComplexNumber z)
    {
        return Divide(Cos(z), Sin(z));
    }

    /// <summary>
    /// Returns the hyperbolic sine of a complex number.
    /// </summary>
    public static ComplexNumber Sinh(ComplexNumber z)
    {
        var real = System.Math.Sinh(z.Real) * System.Math.Cos(z.Imaginary);
        var imag = System.Math.Cosh(z.Real) * System.Math.Sin(z.Imaginary);
        return new ComplexNumber(real, imag);
    }

    /// <summary>
    /// Returns the hyperbolic cosine of a complex number.
    /// </summary>
    public static ComplexNumber Cosh(ComplexNumber z)
    {
        var real = System.Math.Cosh(z.Real) * System.Math.Cos(z.Imaginary);
        var imag = System.Math.Sinh(z.Real) * System.Math.Sin(z.Imaginary);
        return new ComplexNumber(real, imag);
    }

    /// <summary>
    /// Returns the hyperbolic secant of a complex number.
    /// </summary>
    public static ComplexNumber Sech(ComplexNumber z)
    {
        return Divide(new ComplexNumber(1, 0), Cosh(z));
    }

    /// <summary>
    /// Returns the hyperbolic cosecant of a complex number.
    /// </summary>
    public static ComplexNumber Csch(ComplexNumber z)
    {
        return Divide(new ComplexNumber(1, 0), Sinh(z));
    }
}
