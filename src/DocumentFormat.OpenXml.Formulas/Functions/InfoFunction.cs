// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the INFO function.
/// INFO(type_text) - Returns information about the current operating environment.
/// </summary>
/// <remarks>
/// This is a simplified implementation. Full implementation would require access to:
/// - System information (OS version, memory, etc.)
/// - Excel application information
/// - Current calculation settings
///
/// Supported type_text values:
/// "directory" - Path of current directory
/// "numfile" - Number of worksheets in open workbooks
/// "origin" - Cell reference of top-left visible cell (for compatibility)
/// "osversion" - Operating system version
/// "recalc" - Recalculation mode (Automatic/Manual)
/// "release" - Excel version
/// "system" - Operating environment (mac/pcdos)
/// "memavail" - Available memory in bytes
/// "memused" - Memory being used in bytes
/// "totmem" - Total memory available in bytes
/// </remarks>
public sealed class InfoFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly InfoFunction Instance = new();

    private InfoFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "INFO";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 1)
        {
            return FormulaResult.Error("#VALUE!");
        }

        if (args[0].Type != FormulaResultType.Text)
        {
            return FormulaResult.Error("#VALUE!");
        }

        var typeText = args[0].StringValue.ToLowerInvariant();

        return typeText switch
        {
            "directory" => FormulaResult.FromString(Environment.CurrentDirectory),
            "numfile" => FormulaResult.FromNumber(1),
            "origin" => FormulaResult.FromString("$A$1"),
            "osversion" => FormulaResult.FromString(Environment.OSVersion.ToString()),
            "recalc" => FormulaResult.FromString("Automatic"),
            "release" => FormulaResult.FromString("16.0"),
            "system" => GetSystemType(),
            "memavail" => FormulaResult.FromNumber(GetAvailableMemory()),
            "memused" => FormulaResult.FromNumber(GC.GetTotalMemory(false)),
            "totmem" => FormulaResult.FromNumber(GetTotalMemory()),
            _ => FormulaResult.Error("#VALUE!"),
        };
    }

    private static FormulaResult GetSystemType()
    {
#if NET5_0_OR_GREATER
        if (OperatingSystem.IsMacOS())
        {
            return FormulaResult.FromString("mac");
        }
        else if (OperatingSystem.IsWindows())
        {
            return FormulaResult.FromString("pcdos");
        }
        else
        {
            return FormulaResult.FromString("unix");
        }
#else
        // For .NET Standard 2.0, use runtime identifier
        var osDescription = Environment.OSVersion.Platform.ToString().ToLowerInvariant();
        if (osDescription.Contains("unix") || osDescription.Contains("linux"))
        {
            return FormulaResult.FromString("unix");
        }
        else if (osDescription.Contains("win"))
        {
            return FormulaResult.FromString("pcdos");
        }
        else
        {
            return FormulaResult.FromString("mac");
        }
#endif
    }

    private static double GetAvailableMemory()
    {
#if NET5_0_OR_GREATER
        try
        {
            var gcInfo = GC.GetGCMemoryInfo();
            return gcInfo.TotalAvailableMemoryBytes;
        }
        catch
        {
            // Fallback if GC memory info is not available
            return 1073741824; // 1 GB default
        }
#else
        // .NET Standard 2.0 fallback
        return 1073741824; // 1 GB default
#endif
    }

    private static double GetTotalMemory()
    {
#if NET5_0_OR_GREATER
        try
        {
            var gcInfo = GC.GetGCMemoryInfo();
            return gcInfo.TotalAvailableMemoryBytes;
        }
        catch
        {
            // Fallback if GC memory info is not available
            return 1073741824; // 1 GB default
        }
#else
        // .NET Standard 2.0 fallback
        return 1073741824; // 1 GB default
#endif
    }
}
