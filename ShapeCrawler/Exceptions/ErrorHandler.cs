using System;
using System.IO;
using System.Text;

namespace ShapeCrawler.Exceptions;

internal static class ErrorHandler
{
    private static readonly string LogFilePath = Path.Combine(Path.GetTempPath(), "ShapeCrawler.log");

    internal static void Execute(Action action, string context)
    {
        try
        {
            action.Invoke();
        }
        catch (Exception ex)
        {
            var messageBuilder = new StringBuilder();
            messageBuilder.AppendLine(context);
            messageBuilder.AppendLine(ex.ToString());
            File.WriteAllText(LogFilePath, messageBuilder.ToString());

            var userMessage = $@"An error occurred. The full error message has been written to ""{LogFilePath}"" log file. This should not happen, " +
                              "please report this log content as an issue on GitHub https://github.com/ShapeCrawler/ShapeCrawler/issues). " +
                              "We will fix it.";

            throw new ShapeCrawlerException(userMessage, ex);
        }
    }

    internal static void Execute<T>(Func<T> func, string context, out T result)
    {
        try
        {
            result = func.Invoke();
        }
        catch (Exception ex)
        {
            var messageBuilder = new StringBuilder();
            messageBuilder.AppendLine(context);
            messageBuilder.AppendLine(ex.ToString());
            File.WriteAllText(LogFilePath, messageBuilder.ToString());

            var userMessage = $@"An error occurred. The full error message has been written to ""{LogFilePath}"" log file. This should not happen, " +
                              "please report this log content as an issue on GitHub https://github.com/ShapeCrawler/ShapeCrawler/issues). " +
                              "We will fix it.";

            throw new ShapeCrawlerException(userMessage, ex);
        }
    }

    internal static void Execute<T>(Func<T> func, out T result)
    {
        try
        {
            result = func.Invoke();
        }
        catch (Exception ex)
        {
            var messageBuilder = new StringBuilder();
            messageBuilder.AppendLine(ex.ToString());
            File.WriteAllText(LogFilePath, messageBuilder.ToString());

            var userMessage = $@"An error occurred. The full error message has been written to ""{LogFilePath}"" log file. This should not happen, " +
                              "please report this log content as an issue on GitHub https://github.com/ShapeCrawler/ShapeCrawler/issues). " +
                              "We will fix it.";

            throw new ShapeCrawlerException(userMessage, ex);
        }
    }
}