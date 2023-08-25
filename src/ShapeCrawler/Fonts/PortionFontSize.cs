using ShapeCrawler.Services.Factories;
using ShapeCrawler.Shared;
using ShapeCrawler.Texts;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Fonts;

internal record PortionFontSize : IFontSize
{
    private readonly A.Text aText;

    internal PortionFontSize(A.Text aText)
    {
        this.aText = aText;
    }
    
    public int Size()
    {
        var fontSize = this.aText.Parent!.GetFirstChild<A.RunProperties>()?.FontSize
            ?.Value;
        if (fontSize != null)
        {
            return fontSize.Value / 100;
        }

        return SCConstants.DefaultFontSize;
    }

    public void Update(int points)
    {
        var parent = this.aText.Parent!;
        var aRunPr = parent.GetFirstChild<A.RunProperties>();
        if (aRunPr == null)
        {
            var builder = new ARunPropertiesBuilder();
            aRunPr = builder.Build();
            parent.InsertAt(aRunPr, 0);
        }

        aRunPr.FontSize = points * 100;
    }
}