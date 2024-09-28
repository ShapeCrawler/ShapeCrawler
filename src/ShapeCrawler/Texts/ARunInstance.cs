using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Texts;

internal sealed class ARunBuilder
{
    private readonly A.Run aRun;

    internal ARunBuilder()
    {
        this.aRun = new A.Run();
        var aRunProperties = new A.RunProperties { Language = "en-US", FontSize = 1400, Dirty = false };
        var aText = new A.Text
        {
            Text = string.Empty
        };
        this.aRun.Append(aRunProperties);
        this.aRun.Append(aText);
    }

    internal A.Run Build() => this.aRun;
}