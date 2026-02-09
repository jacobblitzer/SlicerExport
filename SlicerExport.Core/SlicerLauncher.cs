using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;

namespace SlicerExport.Core
{
    public static class SlicerLauncher
    {
        public static bool Launch(SlicerTarget target, string threeMfPath)
        {
            string exe = FindSlicerExecutable(target);
            if (exe == null)
                return false;

            Process.Start(new ProcessStartInfo
            {
                FileName = exe,
                Arguments = "\"" + threeMfPath + "\"",
                UseShellExecute = true
            });
            return true;
        }

        private static string FindSlicerExecutable(SlicerTarget target)
        {
            foreach (string path in GetCandidatePaths(target))
            {
                if (File.Exists(path))
                    return path;
            }
            return null;
        }

        private static List<string> GetCandidatePaths(SlicerTarget target)
        {
            var paths = new List<string>();
            bool isMac = RuntimeInformation.IsOSPlatform(OSPlatform.OSX);

            if (isMac)
            {
                switch (target)
                {
                    case SlicerTarget.PrusaSlicer:
                        paths.Add("/Applications/Original Prusa Drivers/PrusaSlicer.app/Contents/MacOS/PrusaSlicer");
                        paths.Add("/Applications/PrusaSlicer.app/Contents/MacOS/PrusaSlicer");
                        break;
                    case SlicerTarget.OrcaSlicer:
                        paths.Add("/Applications/OrcaSlicer.app/Contents/MacOS/OrcaSlicer");
                        break;
                    case SlicerTarget.BambuStudio:
                        paths.Add("/Applications/BambuStudio.app/Contents/MacOS/BambuStudio");
                        break;
                }
            }
            else
            {
                string pf = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
                switch (target)
                {
                    case SlicerTarget.PrusaSlicer:
                        paths.Add(Path.Combine(pf, "Prusa3D", "PrusaSlicer", "prusa-slicer.exe"));
                        break;
                    case SlicerTarget.OrcaSlicer:
                        paths.Add(Path.Combine(pf, "OrcaSlicer", "orca-slicer.exe"));
                        break;
                    case SlicerTarget.BambuStudio:
                        paths.Add(Path.Combine(pf, "Bambu Studio", "bambu-studio.exe"));
                        break;
                }
            }
            return paths;
        }
    }
}
