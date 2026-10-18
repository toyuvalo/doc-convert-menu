// DocConvertLauncher.cs -- DocConvert.exe, the right-click entry point (v1.3.1)
// Replaces launcher.vbs. A wscript-run VBScript that starts a hidden PowerShell
// window is a classic malware chain, and antivirus flagged the installer for it.
// VBScript is also being removed from Windows. Same logic as the VBS:
//   1. Explorer starts one instance per selected file; a lock file lets only the
//      first one do the work.
//   2. Find the Explorer window whose selection holds the right-clicked path (%1)
//      and collect ALL its selected files (multi-select).
//   3. Write the list as UTF-8 and start doc-convert.ps1 with no console window.
// Build: build.cmd (csc.exe from the in-box .NET Framework 4.x).

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading;

[assembly: AssemblyTitle("Doc Convert")]
[assembly: AssemblyDescription("Doc Convert right-click file converter")]
[assembly: AssemblyProduct("Doc Convert Menu")]
[assembly: AssemblyCompany("DVLCE")]
[assembly: AssemblyCopyright("MIT with Commons Clause")]
[assembly: AssemblyVersion("1.3.1.0")]
[assembly: AssemblyFileVersion("1.3.1.0")]

static class DocConvertLauncher
{
    [STAThread]
    static int Main(string[] args)
    {
        string rightClicked = args.Length > 0 ? args[0] : "";
        string tempDir = Path.GetTempPath();
        string lockFile = Path.Combine(tempDir, "doc_convert.lock");

        try
        {
            if (File.Exists(lockFile) &&
                (DateTime.Now - File.GetLastWriteTime(lockFile)).TotalSeconds < 5)
                return 0;
            File.WriteAllText(lockFile, "");
        }
        catch { return 0; }

        // Let Explorer finish launching all instances.
        Thread.Sleep(400);

        var files = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        try { CollectFromExplorer(rightClicked, files, seen); } catch { }

        // Fallback: the right-clicked file alone if no window matched.
        if (files.Count == 0 && rightClicked != "" && File.Exists(rightClicked))
            files.Add(rightClicked);

        try { File.Delete(lockFile); } catch { }

        if (files.Count == 0) return 0;

        string listFile = Path.Combine(tempDir, "doc_convert_batch.txt");
        File.WriteAllLines(listFile, files.ToArray(), new UTF8Encoding(true));

        string script = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "doc-convert.ps1");
        // -ExecutionPolicy Bypass stays until the scripts are code-signed: the
        // Windows client default policy (Restricted) refuses every script.
        var psi = new ProcessStartInfo("powershell.exe",
            "-NoProfile -ExecutionPolicy Bypass -File \"" + script + "\" -ListFile \"" + listFile + "\"");
        psi.UseShellExecute = false;
        psi.CreateNoWindow = true;
        Process.Start(psi);
        return 0;
    }

    static void CollectFromExplorer(string rightClicked, List<string> files, HashSet<string> seen)
    {
        Type shellType = Type.GetTypeFromProgID("Shell.Application");
        dynamic shell = Activator.CreateInstance(shellType);
        foreach (dynamic wnd in shell.Windows())
        {
            dynamic selected;
            try
            {
                if (wnd == null) continue;
                dynamic doc = wnd.Document;
                if (doc == null) continue;
                selected = doc.SelectedItems();
            }
            catch { continue; }   // not a folder view (e.g. an IE window)
            if (selected == null || selected.Count == 0) continue;

            bool match = false;
            foreach (dynamic item in selected)
            {
                if (string.Equals((string)item.Path, rightClicked, StringComparison.OrdinalIgnoreCase))
                { match = true; break; }
            }
            if (!match) continue;

            foreach (dynamic item in selected)
            {
                string p = item.Path;
                if (seen.Add(p)) files.Add(p);
            }
            return;
        }
    }
}
