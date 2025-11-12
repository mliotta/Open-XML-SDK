// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Helper class for day count convention calculations used in financial functions.
/// </summary>
internal static class DayCountHelper
{
    /// <summary>
    /// Calculates day count fraction based on the specified basis.
    /// </summary>
    /// <param name="startDate">Start date.</param>
    /// <param name="endDate">End date.</param>
    /// <param name="basis">Day count basis (0-4).</param>
    /// <returns>Day count fraction.</returns>
    public static double DayCountFraction(DateTime startDate, DateTime endDate, int basis)
    {
        return basis switch
        {
            0 => Basis0_US30_360(startDate, endDate),
            1 => Basis1_ActualActual(startDate, endDate),
            2 => Basis2_Actual360(startDate, endDate),
            3 => Basis3_Actual365(startDate, endDate),
            4 => Basis4_European30_360(startDate, endDate),
            _ => throw new ArgumentException($"Invalid basis: {basis}"),
        };
    }

    /// <summary>
    /// Returns days between two dates based on the specified basis.
    /// </summary>
    public static double DaysBetween(DateTime startDate, DateTime endDate, int basis)
    {
        return basis switch
        {
            0 => Days30_360(startDate, endDate, false),
            1 => (endDate - startDate).TotalDays,
            2 => (endDate - startDate).TotalDays,
            3 => (endDate - startDate).TotalDays,
            4 => Days30_360(startDate, endDate, true),
            _ => throw new ArgumentException($"Invalid basis: {basis}"),
        };
    }

    /// <summary>
    /// Returns days in year for the specified basis.
    /// </summary>
    public static double DaysInYear(int year, int basis)
    {
        return basis switch
        {
            0 => 360.0,
            1 => DateTime.IsLeapYear(year) ? 366.0 : 365.0,
            2 => 360.0,
            3 => 365.0,
            4 => 360.0,
            _ => throw new ArgumentException($"Invalid basis: {basis}"),
        };
    }

    /// <summary>
    /// Basis 0: US (NASD) 30/360.
    /// </summary>
    private static double Basis0_US30_360(DateTime startDate, DateTime endDate)
    {
        var days = Days30_360(startDate, endDate, false);
        return days / 360.0;
    }

    /// <summary>
    /// Basis 1: Actual/actual.
    /// </summary>
    private static double Basis1_ActualActual(DateTime startDate, DateTime endDate)
    {
        var totalDays = (endDate - startDate).TotalDays;
        var startYear = startDate.Year;
        var endYear = endDate.Year;

        if (startYear == endYear)
        {
            var daysInYear = DateTime.IsLeapYear(startYear) ? 366.0 : 365.0;
            return totalDays / daysInYear;
        }
        else
        {
            double yearFraction = 0.0;
            var currentDate = startDate;

            while (currentDate.Year <= endYear)
            {
                var yearStart = new DateTime(currentDate.Year, 1, 1);
                var yearEnd = new DateTime(currentDate.Year, 12, 31);

                var periodStart = currentDate > yearStart ? currentDate : yearStart;
                var periodEnd = endDate < yearEnd ? endDate : yearEnd;

                if (periodStart <= periodEnd)
                {
                    var daysInThisYear = (periodEnd - periodStart).TotalDays;
                    var totalDaysInYear = DateTime.IsLeapYear(currentDate.Year) ? 366.0 : 365.0;
                    yearFraction += daysInThisYear / totalDaysInYear;
                }

                currentDate = new DateTime(currentDate.Year + 1, 1, 1);
            }

            return yearFraction;
        }
    }

    /// <summary>
    /// Basis 2: Actual/360.
    /// </summary>
    private static double Basis2_Actual360(DateTime startDate, DateTime endDate)
    {
        var totalDays = (endDate - startDate).TotalDays;
        return totalDays / 360.0;
    }

    /// <summary>
    /// Basis 3: Actual/365.
    /// </summary>
    private static double Basis3_Actual365(DateTime startDate, DateTime endDate)
    {
        var totalDays = (endDate - startDate).TotalDays;
        return totalDays / 365.0;
    }

    /// <summary>
    /// Basis 4: European 30/360.
    /// </summary>
    private static double Basis4_European30_360(DateTime startDate, DateTime endDate)
    {
        var days = Days30_360(startDate, endDate, true);
        return days / 360.0;
    }

    /// <summary>
    /// Calculates days using 30/360 convention.
    /// </summary>
    /// <param name="startDate">Start date.</param>
    /// <param name="endDate">End date.</param>
    /// <param name="european">True for European method, false for US method.</param>
    private static double Days30_360(DateTime startDate, DateTime endDate, bool european)
    {
        int d1 = startDate.Day;
        int m1 = startDate.Month;
        int y1 = startDate.Year;
        int d2 = endDate.Day;
        int m2 = endDate.Month;
        int y2 = endDate.Year;

        if (european)
        {
            // European 30E/360
            if (d1 == 31)
            {
                d1 = 30;
            }

            if (d2 == 31)
            {
                d2 = 30;
            }
        }
        else
        {
            // US 30US/360
            if (d1 == 31)
            {
                d1 = 30;
            }

            if (d2 == 31 && d1 >= 30)
            {
                d2 = 30;
            }

            // Handle February
            if (m1 == 2 && IsLastDayOfFebruary(startDate))
            {
                d1 = 30;
            }

            if (m2 == 2 && IsLastDayOfFebruary(endDate))
            {
                d2 = 30;
            }
        }

        return ((y2 - y1) * 360) + ((m2 - m1) * 30) + (d2 - d1);
    }

    /// <summary>
    /// Checks if date is last day of February.
    /// </summary>
    private static bool IsLastDayOfFebruary(DateTime date)
    {
        return date.Month == 2 && date.Day == DateTime.DaysInMonth(date.Year, 2);
    }

    /// <summary>
    /// Finds next coupon date after settlement date.
    /// </summary>
    public static DateTime GetNextCouponDate(DateTime settlement, DateTime maturity, int frequency)
    {
        var monthsPerPeriod = 12 / frequency;
        var current = maturity;

        // Work backwards from maturity to find the coupon schedule
        while (current > settlement)
        {
            var previous = current.AddMonths(-monthsPerPeriod);
            if (previous <= settlement)
            {
                return current;
            }

            current = previous;
        }

        return maturity;
    }

    /// <summary>
    /// Finds previous coupon date before settlement date.
    /// </summary>
    public static DateTime GetPreviousCouponDate(DateTime settlement, DateTime maturity, int frequency)
    {
        var nextCoupon = GetNextCouponDate(settlement, maturity, frequency);
        var monthsPerPeriod = 12 / frequency;
        return nextCoupon.AddMonths(-monthsPerPeriod);
    }

    /// <summary>
    /// Counts number of coupons between settlement and maturity.
    /// </summary>
    public static int CountCoupons(DateTime settlement, DateTime maturity, int frequency)
    {
        int count = 0;
        var monthsPerPeriod = 12 / frequency;
        var current = maturity;

        while (current > settlement)
        {
            count++;
            current = current.AddMonths(-monthsPerPeriod);
        }

        return count;
    }

    /// <summary>
    /// Validates frequency is 1, 2, or 4.
    /// </summary>
    public static bool IsValidFrequency(int frequency)
    {
        return frequency == 1 || frequency == 2 || frequency == 4;
    }

    /// <summary>
    /// Validates basis is 0-4.
    /// </summary>
    public static bool IsValidBasis(int basis)
    {
        return basis >= 0 && basis <= 4;
    }
}
