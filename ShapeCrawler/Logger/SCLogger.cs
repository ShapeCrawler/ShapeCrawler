using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.NetworkInformation;
using System.Text;
using System.Text.Json;
using ShapeCrawler.Logger;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

internal static class SCLogger
{
    private static readonly Lazy<SCLog> Log = new(GetLog);
    private static readonly object LoggingLock = new();

    public static void Send()
    {
        if (!SCSettings.CanCollectLogs)
        {
            return;
        }

        if ((DateTime.Now - Log.Value.SentDate).TotalDays < 1)
        {
            return;
        }

        try
        {
            lock (LoggingLock)
            {
                if ((DateTime.Now - Log.Value.SentDate).TotalDays < 1)
                {
                    return;
                }

                using var httpClient = new HttpClient();
                var json = JsonSerializer.Serialize(Log.Value);
                var content = new StringContent(json, Encoding.UTF8, "application/json");
                httpClient.PostAsync("http://domain/api/statistics", content).GetAwaiter().GetResult();

                Log.Value.Reset();
                var path = Path.Combine(Path.GetTempPath(), "sc.log");
                using var fileStream = File.Open(path, FileMode.Create);
                JsonSerializer.Serialize(fileStream, Log);
            }
        }
        catch (Exception)
        {
            // ignored
        }
    }

    private static SCLog GetLog()
    {
        var logFilePath = Path.Combine(Path.GetTempPath(), "sc.log");
        if (File.Exists(logFilePath))
        {
            if (TryGetLogFromFile(logFilePath, out var log))
            {
                return log!;
            }
        }

        NetworkInterface? firstInterface;
        try
        {
            firstInterface = NetworkInterface.GetAllNetworkInterfaces().FirstOrDefault();
        }
        catch (Exception)
        {
            // ignored
            firstInterface = null;
        }

        if (firstInterface == null)
        {
            return new SCLog();
        }

        var newLogValue = new SCLog();
        newLogValue.UserId = firstInterface.GetPhysicalAddress().ToString();
        return newLogValue;
    }

    private static bool TryGetLogFromFile(string logPath, out SCLog? log)
    {
        try
        {
            using var fileStream = File.OpenRead(logPath);
            var fileLogValue = JsonSerializer.Deserialize<SCLog>(fileStream);
            if (fileLogValue != null)
            {
                {
                    log = fileLogValue;
                    return true;
                }
            }
        }
        catch (Exception)
        {
            // ignored
        }

        log = null;
        return false;
    }
}