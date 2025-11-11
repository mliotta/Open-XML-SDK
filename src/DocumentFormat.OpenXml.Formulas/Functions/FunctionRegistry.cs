// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Registry for built-in Excel functions.
/// </summary>
public static class FunctionRegistry
{
    private static readonly Dictionary<string, IFunctionImplementation> _functions = new(StringComparer.OrdinalIgnoreCase)
    {
        // Phase 0 (3)
        { "SUM", SumFunction.Instance },
        { "AVERAGE", AverageFunction.Instance },
        { "IF", IfFunction.Instance },

        // Math (19)
        { "COUNT", CountFunction.Instance },
        { "COUNTA", CountAFunction.Instance },
        { "COUNTBLANK", CountBlankFunction.Instance },
        { "COUNTIF", CountIfFunction.Instance },
        { "MAX", MaxFunction.Instance },
        { "MIN", MinFunction.Instance },
        { "ROUND", RoundFunction.Instance },
        { "ROUNDUP", RoundUpFunction.Instance },
        { "ROUNDDOWN", RoundDownFunction.Instance },
        { "ABS", AbsFunction.Instance },
        { "PRODUCT", ProductFunction.Instance },
        { "POWER", PowerFunction.Instance },
        { "SUMIF", SumIfFunction.Instance },
        { "SQRT", SqrtFunction.Instance },
        { "MOD", ModFunction.Instance },
        { "INT", IntFunction.Instance },
        { "CEILING", CeilingFunction.Instance },
        { "FLOOR", FloorFunction.Instance },
        { "TRUNC", TruncFunction.Instance },

        // Logical (4)
        { "AND", AndFunction.Instance },
        { "OR", OrFunction.Instance },
        { "NOT", NotFunction.Instance },
        { "CHOOSE", ChooseFunction.Instance },

        // Text (14)
        { "CONCATENATE", ConcatenateFunction.Instance },
        { "LEFT", LeftFunction.Instance },
        { "RIGHT", RightFunction.Instance },
        { "MID", MidFunction.Instance },
        { "LEN", LenFunction.Instance },
        { "TRIM", TrimFunction.Instance },
        { "UPPER", UpperFunction.Instance },
        { "LOWER", LowerFunction.Instance },
        { "PROPER", ProperFunction.Instance },
        { "TEXT", TextFunction.Instance },
        { "VALUE", ValueFunction.Instance },
        { "FIND", FindFunction.Instance },
        { "SEARCH", SearchFunction.Instance },
        { "SUBSTITUTE", SubstituteFunction.Instance },

        // Lookup (4)
        { "VLOOKUP", VLookupFunction.Instance },
        { "HLOOKUP", HLookupFunction.Instance },
        { "INDEX", IndexFunction.Instance },
        { "MATCH", MatchFunction.Instance },

        // Date/Time (10)
        { "TODAY", TodayFunction.Instance },
        { "NOW", NowFunction.Instance },
        { "DATE", DateFunction.Instance },
        { "YEAR", YearFunction.Instance },
        { "MONTH", MonthFunction.Instance },
        { "DAY", DayFunction.Instance },
        { "HOUR", HourFunction.Instance },
        { "MINUTE", MinuteFunction.Instance },
        { "SECOND", SecondFunction.Instance },
        { "WEEKDAY", WeekdayFunction.Instance },

        // Statistical (6)
        { "MEDIAN", MedianFunction.Instance },
        { "MODE", ModeFunction.Instance },
        { "STDEV", StDevFunction.Instance },
        { "VAR", VarFunction.Instance },
        { "RANK", RankFunction.Instance },
        { "AVERAGEIF", AverageIfFunction.Instance },

        // Information (7)
        { "ISNUMBER", IsNumberFunction.Instance },
        { "ISTEXT", IsTextFunction.Instance },
        { "IFERROR", IFErrorFunction.Instance },
        { "ISERROR", IsErrorFunction.Instance },
        { "ISNA", IsNaFunction.Instance },
        { "ISERR", IsErrFunction.Instance },
        { "ISBLANK", IsBlankFunction.Instance },
    };

    /// <summary>
    /// Gets a function by name.
    /// </summary>
    /// <param name="name">The function name.</param>
    /// <param name="function">The function implementation.</param>
    /// <returns>True if the function was found.</returns>
    public static bool TryGetFunction(string name, out IFunctionImplementation? function)
    {
        return _functions.TryGetValue(name, out function);
    }
}
