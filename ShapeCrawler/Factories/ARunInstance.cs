using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Factories;

internal class ARunBuilder
{
    private readonly A.Run aRun;

    internal ARunBuilder()
    {
        this.aRun = new A.Run();
        var aRunPropertiesBuilder = new ARunPropertiesBuilder();
        var aRunProperties = aRunPropertiesBuilder.Build();
        var aText = new A.Text
        {
            Text = string.Empty
        };
        this.aRun.Append(aRunProperties);
        this.aRun.Append(aText);
    }

    internal A.Run Build()
    {
        return this.aRun;
    }
}