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
    private static readonly Lazy<Log> Log = new(GetLog);
    private static readonly object LoggingLock = new();
    private static readonly string LogPath = Path.Combine(Path.GetTempPath(), "sc-log.json");

    public static void Send()
    {
        return;

        if (!SCSettings.CanCollectLogs || DateTime.UtcNow < new DateTime(2023, 02, 15))
        {
            return;
        }

        var shouldSend = Log.Value.SentDate == null || (DateTime.Now - Log.Value.SentDate).Value.TotalDays > 1;
        if (!shouldSend)
        {
            return;
        }

        try
        {
            lock (LoggingLock)
            {
                shouldSend = Log.Value.SentDate == null || (DateTime.Now - Log.Value.SentDate).Value.TotalDays > 1; 
                if (!shouldSend)
                {
                    return;
                }

                using var httpClient = new HttpClient();
                var json = JsonSerializer.Serialize(Log.Value);
                var content = new StringContent(json, Encoding.UTF8, "application/json");
                httpClient.PostAsync("http://domain/api/statistics", content).GetAwaiter().GetResult();

                Log.Value.Reset();
                using var fileStream = File.Open(LogPath, FileMode.Create);
                JsonSerializer.Serialize(fileStream, Log.Value);
            }
        }
        catch (Exception)
        {
            Log.Value.SendFailed = DateTime.Now;
        }
    }

    private static Log GetLog()
    {
        if (File.Exists(LogPath))
        {
            if (TryGetLogFromFile(LogPath, out var log))
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
            return new Log();
        }

        var newLogValue = new Log();
        newLogValue.UserId = firstInterface.GetPhysicalAddress().ToString();
        
        return newLogValue;
    }

    private static bool TryGetLogFromFile(string logPath, out Log? log)
    {
        try
        {
            using var fileStream = File.OpenRead(logPath);
            var fileLogValue = JsonSerializer.Deserialize<Log>(fileStream);
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