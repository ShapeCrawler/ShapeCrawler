using System;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Extensions;
using ShapeCrawler.Fonts;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Texts;

internal sealed class TextParagraphPortion : IParagraphPortion
{
    private readonly Lazy<TextPortionFont> font;
    private readonly Lazy<Hyperlink> hyperlink;
    private readonly A.Run aRun;

    internal TextParagraphPortion(OpenXmlPart sdkTypedOpenXmlPart, A.Run aRun)
    {
        this.AText = aRun.Text!;
        this.aRun = aRun;
        var textPortionSize = new PortionFontSize(sdkTypedOpenXmlPart, this.AText);
        this.font = new Lazy<TextPortionFont>(() =>
            new TextPortionFont(sdkTypedOpenXmlPart, this.AText, textPortionSize));
        this.hyperlink = new Lazy<Hyperlink>(() => new Hyperlink(this.aRun.RunProperties!));
    }

    public string Text
    {
        get => this.AText.Text;
        set => this.AText.Text = value;
    }

    public ITextPortionFont Font => this.font.Value;

    public IHyperlink Link => this.hyperlink.Value;

    public Color TextHighlightColor
    {
        get => this.GetTextHighlight();
        set => this.SetTextHighlight(value);
    }

    internal A.Text AText { get; }

    public void Remove() => this.aRun.Remove();

    private Color GetTextHighlight()
    {
        var arPr = this.AText.PreviousSibling<A.RunProperties>();

        // Ensure RgbColorModelHex exists and his value is not null.
        if (arPr?.GetFirstChild<A.Highlight>()?.RgbColorModelHex is not A.RgbColorModelHex aSrgbClr
            || aSrgbClr.Val is null)
        {
            return Color.NoColor;
        }

        // TODO: Check if DocumentFormat.OpenXml.StringValue is necessary.
        var hex = aSrgbClr.Val.ToString() !;

        var color = Color.FromHex(hex);

        var aAlphaValue = aSrgbClr.GetFirstChild<A.Alpha>()?.Val ?? 100000;
        color.Alpha = Color.Opacity / (100000 / aAlphaValue);

        return color;
    }

    private void SetTextHighlight(Color color)
    {
        var arPr = this.AText.PreviousSibling<A.RunProperties>() ?? this.AText.Parent!.AddRunProperties();

        arPr.AddAHighlight(color);
    }
}