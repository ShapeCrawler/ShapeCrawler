namespace ShapeCrawler.Factories
{
    internal static class ARunInstance
    {
        internal static DocumentFormat.OpenXml.Drawing.Run CreateEmpty()
        {
            var aRun = new DocumentFormat.OpenXml.Drawing.Run();
            var aRunProperties = new DocumentFormat.OpenXml.Drawing.RunProperties { Language = "en-US", FontSize = 1400, Dirty = false };
            var aText = new DocumentFormat.OpenXml.Drawing.Text
            {
                Text = string.Empty
            };
            aRun.Append(aRunProperties);
            aRun.Append(aText);

            return aRun;
        }
    }
}