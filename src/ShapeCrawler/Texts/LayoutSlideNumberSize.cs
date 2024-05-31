using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Fonts;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Texts;

internal record LayoutSlideNumberSize : IFontSize
{
    private readonly P.TextBody pTextBody;
    private readonly ISlideNumberFont masterSlideNumberFont;
    private const decimal HalfPointsInPoint = 100m;

    internal LayoutSlideNumberSize(P.TextBody pTextBody, ISlideNumberFont masterSlideNumberFont)
    {
        this.pTextBody = pTextBody;
        this.masterSlideNumberFont = masterSlideNumberFont;
    }

    public decimal Size()
    {
        var halfPoints = this.pTextBody.Descendants<A.Field>().First().RunProperties?.FontSize?.Value;
        if (halfPoints == null)
        {
            return this.masterSlideNumberFont.Size;
        }

        var points = halfPoints.Value / HalfPointsInPoint;

        return points;
    }

    public void Update(decimal points)
    {
        var aListStyle = this.pTextBody.ListStyle!;
        var aLvl1pPr = aListStyle.Level1ParagraphProperties;
        aLvl1pPr?.Remove();

        var halfPoints = points * HalfPointsInPoint;
        aListStyle.AppendChild(
            new A.Level1ParagraphProperties(
                new A.DefaultRunProperties { FontSize = new Int32Value((int)halfPoints) }));
    }
}