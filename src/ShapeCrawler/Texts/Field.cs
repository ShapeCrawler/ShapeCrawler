using System;
using ShapeCrawler.Extensions;
using ShapeCrawler.Fonts;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Texts;

internal sealed class Field : IParagraphPortion
{
    private readonly Lazy<ITextPortionFont> font;
    private readonly Lazy<Hyperlink> hyperlink;
    private readonly A.Field aField;
    private readonly FieldPortionText fieldPortionText;
    private readonly A.Text? aText;

    internal Field(A.Field aField)
    {
        this.aText = aField.GetFirstChild<A.Text>();
        this.aField = aField;
        this.font = new Lazy<ITextPortionFont>(() =>
        {
            var textPortionSize = new PortionFontSize(this.aText!);
            return new TextPortionFont(this.aText!, textPortionSize);
        });
        this.fieldPortionText = new FieldPortionText(this.aField);
        this.hyperlink = new Lazy<Hyperlink>(() => new Hyperlink(this.aField.RunProperties!));
    }

    /// <inheritdoc/>
    public string Text
    {
        get => this.fieldPortionText.Value;
        set => this.fieldPortionText.Update(value);
    }

    /// <inheritdoc/>
    public ITextPortionFont Font => this.font.Value;

    public IHyperlink Link => this.hyperlink.Value;

    public Color TextHighlightColor
    {
        get => this.ParseTextHighlight();
        set => this.UpdateTextHighlight(value);
    }

    public void Remove() => this.aField.Remove();

    private Color ParseTextHighlight()
    {
        var arPr = this.aText!.PreviousSibling<A.RunProperties>();

        // Ensure RgbColorModelHex exists and his value is not null.
        if (arPr?.GetFirstChild<A.Highlight>()?.RgbColorModelHex is not A.RgbColorModelHex aSrgbClr
            || aSrgbClr.Val is null)
        {
            return Color.NoColor;
        }

        // Gets node value.
        // TODO: Check if DocumentFormat.OpenXml.StringValue is necessary.
        var hex = aSrgbClr.Val.ToString() !;

        // Check if color value is valid, we are expecting values as "000000".
        var color = Color.FromHex(hex);

        // Calculate alpha value if is defined in highlight node.
        var aAlphaValue = aSrgbClr.GetFirstChild<A.Alpha>()?.Val ?? 100000;
        color.Alpha = Color.Opacity / (100000 / aAlphaValue);

        return color;
    }

    private void UpdateTextHighlight(Color color)
    {        
        var arPr = this.aText!.PreviousSibling<A.RunProperties>() ?? this.aText.Parent!.AddRunProperties();

        arPr.AddAHighlight(color);
    }
}