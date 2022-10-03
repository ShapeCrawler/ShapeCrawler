using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Factories;

internal static class ARunInstance
{
    internal static A.Run? CreateEmpty()
    {
        var aRun = new A.Run();
        var aRunPropertiesBuilder = new ARunPropertiesBuilder();
        var aRunProperties = aRunPropertiesBuilder.Build();
        var aText = new A.Text
        {
            Text = string.Empty
        };
        aRun.Append(aRunProperties);
        aRun.Append(aText);

        return aRun;
    }
}