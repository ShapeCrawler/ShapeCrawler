using System;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using ShapeCrawler.Shapes;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler.Texts;

internal sealed class TextBox : ITextBox
{
    private readonly TextBoxMargins margins;
    private readonly OpenXmlElement textBody;
    private readonly ShapeSize shapeSize;
    private readonly TextAutofit autofit;
    private TextVerticalAlignment? vAlignment;
    private TextDirection? textDirection;

    internal TextBox(TextBoxMargins margins, OpenXmlElement textBody)
    {
        this.margins = margins;
        this.textBody = textBody;
        this.shapeSize = new ShapeSize(textBody.Parent!);
        this.autofit = new TextAutofit(
            this.Paragraphs,
            () => this.AutofitType,
            this.shapeSize,
            this.margins,
            () => this.TextWrapped,
            this.textBody);
    }

    public IParagraphCollection Paragraphs => new ParagraphCollection(this.textBody);

    public string Text
    {
        get
        {
            var stringBuilder = new StringBuilder();
            stringBuilder.Append(this.Paragraphs[0].Text);

            var paragraphsCount = this.Paragraphs.Count;
            var index = 1; // we've already added the text of first paragraph
            while (index < paragraphsCount)
            {
                stringBuilder.AppendLine();
                stringBuilder.Append(this.Paragraphs[index].Text);

                index++;
            }

            return stringBuilder.ToString();
        }
    }

    public AutofitType AutofitType
    {
        get
        {
            var aBodyPr = this.textBody.GetFirstChild<A.BodyProperties>();

            if (aBodyPr!.GetFirstChild<A.NormalAutoFit>() != null)
            {
                return AutofitType.Shrink;
            }

            if (aBodyPr.GetFirstChild<A.ShapeAutoFit>() != null)
            {
                return AutofitType.Resize;
            }

            return AutofitType.None;
        }
        
        set
        {
            var currentType = this.AutofitType;
            if (currentType == value)
            {
                return;
            }

            var aBodyPr = this.textBody.GetFirstChild<A.BodyProperties>() !;

            RemoveExistingAutofitElements(aBodyPr);

            switch (value)
            {
                case AutofitType.None:
                    aBodyPr.Append(new A.NoAutoFit());
                    break;
                case AutofitType.Shrink:
                    aBodyPr.Append(new A.NormalAutoFit());
                    break;
                case AutofitType.Resize:
                    aBodyPr.Append(new A.ShapeAutoFit());
                    this.ResizeParentShapeOnDemand();
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }

    public decimal LeftMargin
    {
        get => this.margins.Left;
        set => this.margins.Left = value;
    }

    public decimal RightMargin
    {
        get => this.margins.Right;
        set => this.margins.Right = value;
    }

    public decimal TopMargin
    {
        get => this.margins.Top;
        set
        {
            this.margins.Top = value;
        }
    }

    public decimal BottomMargin
    {
        get => this.margins.Bottom;
        set => this.margins.Bottom = value;
    }

    public bool TextWrapped
    {
        get
        {
            var aBodyPr = this.textBody.GetFirstChild<A.BodyProperties>() !;
            var wrap = aBodyPr.GetAttributes().FirstOrDefault(a => a.LocalName == "wrap");

            if (wrap.Value == "none")
            {
                return false;
            }

            return true;
        }
    }

    public string SDKXPath => new XmlPath(this.textBody).XPath;

    public TextVerticalAlignment VerticalAlignment
    {
        get
        {
            if (this.vAlignment.HasValue)
            {
                return this.vAlignment.Value;
            }

            var aBodyPr = this.textBody.GetFirstChild<A.BodyProperties>();

            if (aBodyPr!.Anchor?.Value == A.TextAnchoringTypeValues.Center)
            {
                this.vAlignment = TextVerticalAlignment.Middle;
            }
            else if (aBodyPr.Anchor?.Value == A.TextAnchoringTypeValues.Bottom)
            {
                this.vAlignment = TextVerticalAlignment.Bottom;
            }
            else
            {
                this.vAlignment = TextVerticalAlignment.Top;
            }

            return this.vAlignment.Value;
        }

        set => this.SetVerticalAlignment(value);
    }

    public TextDirection TextDirection 
    {
        get
        {
            if (!this.textDirection.HasValue)
            {
                var textDirectionVal = this.textBody.GetFirstChild<A.BodyProperties>()!.Vertical?.Value;

                if (textDirectionVal == A.TextVerticalValues.Vertical)
                {
                    this.textDirection = TextDirection.Rotate90;
                }
                else if (textDirectionVal == A.TextVerticalValues.Vertical270)
                {
                    this.textDirection = TextDirection.Rotate270;
                }
                else if (textDirectionVal == A.TextVerticalValues.WordArtVertical)
                {
                    this.textDirection = TextDirection.Stacked;
                }
                else
                {
                    this.textDirection = TextDirection.Horizontal;
                }
            }

            return this.textDirection.Value;
        }

        set => this.SetTextDirection(value); 
    }

    public void SetMarkdownText(string text)
    {
        var markdownText = new MarkdownText(
            text,
            this.Paragraphs,
            () => this.AutofitType,
            this.autofit.ShrinkFont,
            this.autofit.Apply);
        markdownText.ApplyTo();
    }

    public void SetText(string text)
    {
        var textContent = new TextContent(
            text,
            this.Paragraphs,
            () => this.AutofitType,
            this.autofit.ShrinkFont,
            this.autofit.Apply);
        textContent.ApplyTo();
    }

    internal void ResizeParentShapeOnDemand()
    {
        this.autofit.Apply();
    }

    private static void RemoveExistingAutofitElements(A.BodyProperties bodyProperties)
    {
        bodyProperties.GetFirstChild<A.NoAutoFit>()?.Remove();
        bodyProperties.GetFirstChild<A.NormalAutoFit>()?.Remove();
        bodyProperties.GetFirstChild<A.ShapeAutoFit>()?.Remove();
    }

    private void SetVerticalAlignment(TextVerticalAlignment alignmentValue)
    {
        var aTextAlignmentTypeValue = alignmentValue switch
        {
            TextVerticalAlignment.Top => A.TextAnchoringTypeValues.Top,
            TextVerticalAlignment.Middle => A.TextAnchoringTypeValues.Center,
            TextVerticalAlignment.Bottom => A.TextAnchoringTypeValues.Bottom,
            _ => throw new ArgumentOutOfRangeException(nameof(alignmentValue))
        };

        var aBodyPr = this.textBody.GetFirstChild<A.BodyProperties>();

        aBodyPr?.Anchor = aTextAlignmentTypeValue;

        this.vAlignment = alignmentValue;
    }

    private void SetTextDirection(TextDirection direction)
    {
        var aBodyPr = this.textBody.GetFirstChild<A.BodyProperties>()!;
         
        aBodyPr.Vertical = direction switch
        {
            TextDirection.Rotate90 => A.TextVerticalValues.Vertical,
            TextDirection.Rotate270 => A.TextVerticalValues.Vertical270,
            TextDirection.Stacked => A.TextVerticalValues.WordArtVertical,
            _ => A.TextVerticalValues.Horizontal
        };

        this.textDirection = direction;
    }

}