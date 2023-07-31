using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Services.Factories;

internal sealed class ARunPropertiesBuilder
{
    private readonly A.RunProperties aRunProperties = new () { Language = "en-US", FontSize = 1400, Dirty = false };

    internal A.RunProperties Build()
    {
        return this.aRunProperties;
    }
}