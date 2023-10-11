using ShapeCrawler.Services.Factories;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Fonts;

internal class PortionFontSize : IFontSize
{
    private readonly A.Text aText;

    internal PortionFontSize(A.Text aText)
    {
        this.aText = aText;
    }
    
    int IFontSize.Size()
    {
        var fontSize = this.aText.Parent!.GetFirstChild<A.RunProperties>()?.FontSize
            ?.Value;
        if (fontSize != null)
        {
            return fontSize.Value / 100;
        }

        return Constants.DefaultFontSize;
    }

    void IFontSize.Update(int points)
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