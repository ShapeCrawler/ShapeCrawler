using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Paragraphs;
using ShapeCrawler.Shapes;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Fonts;

internal class PortionFontSize(OpenXmlPart openXmlPart, A.Text aText): IFontSize
{
    decimal IFontSize.Size
    {
        get
        {
            var hundredsPoints = aText.Parent!.GetFirstChild<A.RunProperties>()?.FontSize
                ?.Value;
            if (hundredsPoints != null)
            {
                return hundredsPoints.Value / 100m;
            }

            hundredsPoints = new ReferencedIndentLevel(openXmlPart, aText).FontSizeOrNull();
            if (hundredsPoints is not null)
            {
                return hundredsPoints.Value / 100m;
            }

            var indentLevel = new SCAParagraph(aText.Ancestors<A.Paragraph>().First()).GetIndentLevel();
            SlideMasterPart slideMasterPart;
            if (openXmlPart is SlideMasterPart)
            {
                slideMasterPart = (SlideMasterPart)openXmlPart;
            }
            else
            {
                slideMasterPart = ((SlidePart)openXmlPart).SlideLayoutPart!.SlideMasterPart!;
            }

            AutoShape? parentAutoShape = null;
            var parentPShape = aText.Ancestors<P.Shape>().FirstOrDefault();
            if (parentPShape is not null)
            {
                parentAutoShape = new AutoShape(openXmlPart, aText.Ancestors<P.Shape>().First());
            }

            if (parentAutoShape is not null && parentAutoShape.IsPlaceholder)
            {
                if (parentAutoShape.PlaceholderType == PlaceholderType.Title)
                {
                    var titleFontSizeHundredPoints =
                        slideMasterPart.SlideMaster.TextStyles!.TitleStyle!.Level1ParagraphProperties!
                            .GetFirstChild<A.DefaultRunProperties>() !.FontSize!.Value;

                    return titleFontSizeHundredPoints / 100m;
                }

                var indentFonts =
                    new IndentFonts(slideMasterPart.SlideMaster.TextStyles!.BodyStyle!);
                var indentFont = indentFonts.FontOrNull(indentLevel);
                if (indentFont != null)
                {
                    return indentFont.Value.Size!.Value / 100m;
                }
            }

            // Presentation
            var pPresentation = ((PresentationDocument)openXmlPart.OpenXmlPackage).PresentationPart!
                .Presentation;
            if (pPresentation.DefaultTextStyle != null)
            {
                var defaultTextStyleFonts = new IndentFonts(pPresentation.DefaultTextStyle);
                var defaultTextStyleFont = defaultTextStyleFonts.FontOrNull(indentLevel);
                if (defaultTextStyleFont.HasValue && defaultTextStyleFont.Value.Size != null)
                {
                    return defaultTextStyleFont.Value.Size!.Value / 100m;
                }

                var aTextDefault2 = pPresentation.PresentationPart!.ThemePart!.Theme.ObjectDefaults!.TextDefault;
                if (aTextDefault2 is not null)
                {
                    var listStyleFonts =
                        new IndentFonts(pPresentation.PresentationPart!.ThemePart!.Theme.ObjectDefaults!.TextDefault!
                            .ListStyle!);
                    var listStyleFontsFont = listStyleFonts.FontOrNull(indentLevel);
                    if (listStyleFontsFont.HasValue && listStyleFontsFont.Value.Size != null)
                    {
                        return listStyleFontsFont.Value.Size!.Value / 100m;
                    }
                }
            }

            if (parentAutoShape is not null && parentAutoShape.IsPlaceholder)
            {
                var indentFonts2 =
                    new IndentFonts(slideMasterPart.SlideMaster.TextStyles!.BodyStyle!);
                var indentFont2 = indentFonts2.FontOrNull(indentLevel);
                if (indentFont2 != null && indentFont2.Value.Size != null)
                {
                    return indentFont2.Value.Size!.Value / 100m;
                }
            }

            var aTextDefault = pPresentation.PresentationPart!.ThemePart!.Theme.ObjectDefaults!.TextDefault;
            if (aTextDefault is not null)
            {
                var listStyleFonts = new IndentFonts(aTextDefault.ListStyle!);
                var listStyleFont = listStyleFonts.FontOrNull(indentLevel);
                if (listStyleFont.HasValue && listStyleFont.Value.Size != null)
                {
                    return listStyleFont.Value.Size!.Value / 100m;
                }
            }

            return 18; // default: https://bit.ly/37Tjjlo
        }

        set
        {
            var parent = aText.Parent!;
            var aRunPr = parent.GetFirstChild<A.RunProperties>();
            if (aRunPr == null)
            {
                aRunPr = new A.RunProperties { Language = "en-US", FontSize = 1400, Dirty = false };
                parent.InsertAt(aRunPr, 0);
            }
            
            aRunPr.FontSize = (int)(value * 100);
        }
    }
}