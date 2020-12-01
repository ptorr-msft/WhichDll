# WhichDll
A simple tool to find which DLL is housing an API exported by a LIB. 

Output from the tool:

```
WhichDll: Which DLL exports a given function, according to an implib?
          Source available at https://github.com/ptorr-msft/WhichDll.

Usage: WhichDll <implib> <export-prefix> [-nologo]

For example, to find out which DLL contains CreateFileFromAppW according
to OneCoreUap.lib, you can use any of the following:

       WhichDll onecoreuap.lib CreateFileFromAppW
       WhichDll onecoreuap CreateFileFrom
       WhichDll onecoreuap createfile

The export name is case-insensitive and will report all functions that match
the given prefix. So, for example, the third command-line above will return
results for other exports such as CreateFileA, CreateFile2, etc.

If the specified DLL is actually an API Set, WhichDll will attempt to locate
the actual DLL that hosts the API _on_this_machine_; please note that it
could resolve to a different DLL on a different machine, so you should not
depend on this information for anything other than local debugging.

If you specify '-i' as the <implib>, WhichDll will read from stdin. Useful
for piping the output of dumpbin (or something else) into the app:

       c:\path\to\dumpbin -all c:\path\to\foo.lib | whichdll -i someexport

The '-nologo' switch hides the banner and other non-essential output.
```
