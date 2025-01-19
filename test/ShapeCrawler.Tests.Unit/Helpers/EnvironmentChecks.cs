namespace ShapeCrawler.Tests.Unit.Helpers;

public static class EnvironmentChecks
{
    public static bool IsGhostscriptInstalled()
    {
        try
        {
            // Try to locate Ghostscript executable
            var processStartInfo = new System.Diagnostics.ProcessStartInfo
            {
                FileName = "gs", // 'gs' is the typical command for Ghostscript
                Arguments = "--version", // Check Ghostscript version as an example
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using (var process = System.Diagnostics.Process.Start(processStartInfo))
            {
                process.WaitForExit();
                return process.ExitCode == 0;
            }
        }
        catch
        {
            return false; // Assume not installed if there's an exception
        }
    }
}