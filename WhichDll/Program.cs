﻿using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace WhichDll
{
    static class Win32Interop
    {
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        static extern IntPtr LoadLibraryW([MarshalAs(UnmanagedType.LPWStr)] string libraryName);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        static extern Int32 GetModuleFileNameW(IntPtr handle, StringBuilder filename, Int32 size);

        const Int32 MAX_PATH = 260;

        internal static bool IsApiSetName(string dllName)
        {
            return dllName.StartsWith("api-") || dllName.StartsWith("ext-");
        }

        internal static string GetRealModuleFilename(string apisetName)
        {
            var handle = LoadLibraryW(apisetName);
            if (handle != IntPtr.Zero)
            {
                StringBuilder realName = new StringBuilder(MAX_PATH);

                int result = GetModuleFileNameW(handle, realName, MAX_PATH);
                if (result > 0)
                {
                    return realName.ToString();
                }
            }

            return apisetName;
        }
    }

    class Program
    {
        enum ReturnCode
        {
            Success = 0,
            ExportNotFound,
            BadCommandLine,
            DumpbinNotFound,
            LibNotFound,
            NotALibFile,
            CantOpenFile,
            OtherError,
        }

        static bool showNonEssentialOutput = true;

        static int Main(string[] args)
        {
            int argsWithoutNologo = args.Length;

            if (args.Length > 0 && args[args.Length - 1].ToLowerInvariant().Substring(1) == "nologo")
            {
                showNonEssentialOutput = false;
                --argsWithoutNologo;
            }
            else
            {
                Console.WriteLine(@"WhichDll: Which DLL exports a given function, according to an implib?
          Source available at https://github.com/ptorr-msft/WhichDll.
");
            }

            if (argsWithoutNologo < 1)
            {
                Usage();
                return (int)ReturnCode.BadCommandLine;
            }

            var exportQuery = args[0];
            string libFilename = "onecoreuap.lib";
            if (argsWithoutNologo >= 2)
            {
                libFilename = args[1];
            }

            TextReader dumpReader = null;
            var result = ReturnCode.OtherError;
            string displayFilename;

            // Special filename means "read from stdin"
            if (libFilename == "-i" || libFilename == "/i")
            {
                displayFilename = "<stdin>";
                dumpReader = new StreamReader(Console.OpenStandardInput());
            }
            else
            {
                displayFilename = Path.GetFileName(libFilename);
                try
                {
                    dumpReader = GetOutputStreamFromDumpbin(libFilename);
                }
                catch (Win32Exception)
                {
                    Console.Error.WriteLine("Can't find dumpbin.exe; is it in your PATH?");
                    result = ReturnCode.DumpbinNotFound;
                }
            }

            if (dumpReader != null)
            {
                result = ProcessDumpbinOutput(displayFilename, dumpReader, exportQuery);
            }

            return (int)result;
        }

        private static TextReader GetOutputStreamFromDumpbin(string libFilename)
        {
            if (!libFilename.Contains("."))
            {
                libFilename += ".lib";
            }

            var exeName = FindDumpbin();
            var parameters = $@"-headers ""{libFilename}""";

            if (showNonEssentialOutput)
            {
                Console.WriteLine($@"Executing '{exeName} {parameters}'...");
                Console.WriteLine();
            }

            var dumpbinStartInfo = new ProcessStartInfo(exeName, parameters);
            dumpbinStartInfo.UseShellExecute = false;
            dumpbinStartInfo.RedirectStandardOutput = true;

            var dumpbinProcess = Process.Start(dumpbinStartInfo);

            return dumpbinProcess.StandardOutput;
        }

        private static string FindDumpbin()
        {
            if (Debugger.IsAttached)
            {
                // Keep up to date with whatever SDK is installed...
                return @"C:\Program Files (x86)\Microsoft Visual Studio\2019\Preview\VC\Tools\MSVC\14.28.29333\bin\Hostx64\x64\dumpbin.exe";
            }

            var fullPath = "dumpbin.exe";
            var whereStartInfo = new ProcessStartInfo("where", "dumpbin.exe");
            whereStartInfo.UseShellExecute = false;
            whereStartInfo.RedirectStandardOutput = true;

            try
            {
                var whereProcess = Process.Start(whereStartInfo);
                fullPath = whereProcess.StandardOutput.ReadLine() ?? fullPath;
            }
            catch
            {
                // Meh... will fail later anyway.
            }

            return fullPath;
        }

        static ReturnCode ProcessDumpbinOutput(string displayFilename, TextReader output, string exportQuery)
        {
            string line;
            string exportingDll = string.Empty;
            string previousLine = string.Empty;
            string fullExportName;
            int howManyFound = 0;

            try
            {
                while (null != (line = output.ReadLine()))
                {
                    Debug.WriteLine(line);

                    if (line.Length == 0)
                    {
                        continue;
                    }

                    ReturnCode returnCode;
                    if (TryGetDumpbinError(line, out returnCode))
                    {
                        ProcessDumpbinError(displayFilename, line, returnCode);
                        return returnCode;
                    }

                    if (TryGetFullExportName(line, exportQuery, out fullExportName))
                    {
                        if (TryGetDllName(previousLine, ref exportingDll))
                        {
                            ++howManyFound;
                            if (Win32Interop.IsApiSetName(exportingDll))
                            {
                                var realFilename = Win32Interop.GetRealModuleFilename(exportingDll);
                                if (realFilename != exportingDll)
                                {
                                    exportingDll += $" --> {Path.GetFileName(realFilename)} on this machine";
                                }
                            }

                            Console.WriteLine($"{fullExportName} is exported by {exportingDll.ToLowerInvariant()}.");
                        }
                    }

                    previousLine = line;
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex}");
                return ReturnCode.OtherError;
            }

            if (howManyFound > 0)
            {
                if (showNonEssentialOutput)
                {
                    Console.WriteLine();
                    Console.WriteLine($"Found {howManyFound} export(s) matching '{exportQuery}' from {displayFilename}.");
                }
                return ReturnCode.Success;
            }
            else
            {
                if (showNonEssentialOutput)
                {
                    Console.WriteLine($"No exports found matching '{exportQuery}' from {displayFilename}.");
                }
                return ReturnCode.ExportNotFound;
            }
        }

        private static void ProcessDumpbinError(string displayFilename, string line, ReturnCode returnCode)
        {
            switch (returnCode)
            {
                case ReturnCode.LibNotFound:
                    Console.Error.WriteLine($"'{displayFilename}' cannot be found; is the LIB environment variable defined?");
                    break;

                case ReturnCode.CantOpenFile:
                    Console.Error.WriteLine($"'{displayFilename}' cannot be opened; do you have access to it?");
                    break;

                case ReturnCode.NotALibFile:
                    Console.Error.WriteLine($"'{displayFilename}' is not a valid PE file.");
                    break;

                default:
                    Console.Error.WriteLine($"Error reported from Dumpbin: {line}");
                    break;
            }
        }

        private static bool TryGetDumpbinError(string line, out ReturnCode error)
        {
            error = ReturnCode.Success;

            if (line.Contains("fatal error LNK1181"))
            {
                error = ReturnCode.LibNotFound;
                return true;
            }
            else if (line.Contains("warning LNK4048"))
            {
                error = ReturnCode.NotALibFile;
                return true;
            }
            else if (line.Contains("fatal error LNK1104"))
            {
                error = ReturnCode.CantOpenFile;
                return true;
            }

            return false;
        }

        static Regex dllNameMatch = new Regex(@"DLL name\s+:\s+(\S+)");
        static Regex exportNameMatch = new Regex(@"Symbol name\s+:\s+(\S+)");
        static Regex isRegexMatch = new Regex(@"[^a-zA-Z0-9\-_]");

        private static bool TryGetFullExportName(string line, string exportQuery, out string fullExportName)
        {
            var match = exportNameMatch.Match(line);
            if (match.Success && match.Groups?.Count > 1)
            {
                var exportName = match.Groups[1].Value;
                if (isRegexMatch.Match(exportQuery).Success)
                {
                    var exportMatch = new Regex(exportQuery);
                    if (exportMatch.Match(exportName).Success)
                    {
                        fullExportName = exportName;
                        return true;
                    }
                }
                else
                {
                    var comparison = String.Compare(exportQuery, 0, exportName, 0, exportQuery.Length, true);
                    if (comparison == 0)
                    {
                        fullExportName = exportName;
                        return true;
                    }
                }
            }

            fullExportName = string.Empty;
            return false;
        }

        private static bool TryGetDllName(string line, ref string dllName)
        {
            var match = dllNameMatch.Match(line);
            if (match.Success && match.Groups?.Count > 1)
            {
                dllName = match.Groups[1].Value;
                return true;
            }

            return false;
        }

        static void Usage()
        {
            Console.WriteLine(@"Usage: WhichDll <export> [<implib>] [-nologo]

By default, the <export> is case-insensitive and will report all functions 
that match the given prefix. If <implib> is not provided, it defaults to
onecoreuap.lib.

For example, to find out which DLL contains CreateFileFromAppW according to 
OneCoreUap.lib, you can use any of the following:

       WhichDll CreateFileFromAppW onecoreuap.lib 
       WhichDll CreateFileFrom onecoreuap 
       WhichDll createfile

Third command-line above will return results for other exports that begin with
'createfile' such as CreateFileA, CreateFile2, etc.

You can also use a regular expression for the <export>. For example, the 
following finds CreateFileFromAppW and a couple of other exports:

       WhichDll Create.*AppW

If the DLL referenced in the implib is actually an API Set, WhichDll will 
attempt to locate the DLL that hosts the API *on this machine*  please note 
that it could resolve to a different DLL on a different machine, so you 
should *not* depend on this information for anything other than local 
debugging.

If you specify '-i' as the <implib>, WhichDll will read from stdin. this can 
be useful for piping the output of dumpbin (or something else) into the app:

       c:\path\to\dumpbin -all c:\path\to\foo.lib | whichdll someexport -i

The -nologo switch hides the banner and other non-essential output.");
        }
    }
}
