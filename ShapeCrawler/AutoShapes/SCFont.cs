using System;
using DocumentFormat.OpenXml;
using ShapeCrawler.Drawing;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Extensions;
using ShapeCrawler.Factories;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Services;
using ShapeCrawler.Shared;
using ShapeCrawler.SlideMasters;
using ShapeCrawler.Statics;
using ShapeCrawler.Tables;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.AutoShapes
{
    internal class SCFont : IFont
    {
        private readonly A.Text aText;
        private readonly A.FontScheme aFontScheme;
        private readonly Lazy<ColorFormat> colorFormat;
        private readonly ResettableLazy<A.LatinFont> latinFont;
        private readonly ResettableLazy<int> size;

        internal SCFont(A.Text aText, SCPortion portion)
        {
            this.aText = aText;
            this.size = new ResettableLazy<int>(this.GetSize);
            this.latinFont = new ResettableLazy<A.LatinFont>(this.GetALatinFont);
            this.colorFormat = new Lazy<ColorFormat>(() => new ColorFormat(this));
            this.ParentPortion = portion;
            var parentTextBoxContainer = portion.ParentParagraph.ParentTextBox.TextFrameContainer;
            Shape parentShape;
            if (parentTextBoxContainer is SCTableCell cell)
            {
                parentShape = (Shape)cell.Shape;
            }
            else
            {
                parentShape = (Shape)portion.ParentParagraph.ParentTextBox.TextFrameContainer;
            }

            this.aFontScheme = parentShape.SlideMasterInternal.ThemePart.Theme.ThemeElements.FontScheme;
        }

        #region Public Properties

        public string Name
        {
            get => this.GetName();
            set => this.SetName(value);
        }

        public int Size
        {
            get => this.size.Value;
            set => this.SetFontSize(value);
        }

        public bool IsBold
        {
            get => this.GetBoldFlag();
            set => this.SetBoldFlag(value);
        }

        public bool IsItalic
        {
            get => this.GetItalicFlag();
            set => this.SetItalicFlag(value);
        }

        public DocumentFormat.OpenXml.Drawing.TextUnderlineValues Underline
        {
            get
            {
                A.RunProperties aRunProperties = this.aText.Parent.GetFirstChild<A.RunProperties>();
                return aRunProperties?.Underline?.Value ?? A.TextUnderlineValues.None;
            }
            set
            {
                A.RunProperties aRunPr = this.aText.Parent.GetFirstChild<A.RunProperties>();
                if (aRunPr != null)
                {
                    aRunPr.Underline = new EnumValue<A.TextUnderlineValues>(value);
                }
                else
                {
                    A.EndParagraphRunProperties aEndParaRPr = this.aText.Parent.NextSibling<A.EndParagraphRunProperties>();
                    if (aEndParaRPr != null)
                    {
                        aEndParaRPr.Underline = new EnumValue<A.TextUnderlineValues>(value);
                    }
                    else
                    {
                        var runProp = this.aText.Parent.AddRunProperties();
                        runProp.Underline = new EnumValue<A.TextUnderlineValues>(value);
                    }
                }
            }
        }

        public IColorFormat ColorFormat => this.colorFormat.Value;

        #endregion Public Properties

        internal SCPortion ParentPortion { get; }

        public bool CanChangeSize()
        {
            A.RunProperties runPr = this.aText.Parent.GetFirstChild<A.RunProperties>();
            return runPr != null;
        }

        private string GetName()
        {
            const string majorLatinFont = "+mj-lt";
            if (this.latinFont.Value.Typeface == majorLatinFont)
            {
                return this.aFontScheme.MajorFont.LatinFont.Typeface;
            }

            return this.latinFont.Value.Typeface;
        }

        private A.LatinFont GetALatinFont()
        {
            A.RunProperties aRunProperties = this.aText.Parent.GetFirstChild<A.RunProperties>();
            A.LatinFont aLatinFont = aRunProperties?.GetFirstChild<A.LatinFont>();

            if (aLatinFont != null)
            {
                return aLatinFont;
            }

            FontData phFontData = new ();
            FontDataParser.GetFontDataFromPlaceholder(ref phFontData, this.ParentPortion.ParentParagraph);
            {
                if (phFontData.ALatinFont != null)
                {
                    return phFontData.ALatinFont;
                }
            }

            // Get from theme
            return this.aFontScheme.MinorFont.LatinFont;
        }

        private int GetSize()
        {
            var fontSize = this.ParentPortion.SDKAText.Parent!.GetFirstChild<A.RunProperties>()?.FontSize?.Value;
            if (fontSize != null)
            {
                return fontSize.Value / 100;
            }

            var parentParagraph = this.ParentPortion.ParentParagraph;
            var textBoxContainer = parentParagraph.ParentTextBox.TextFrameContainer;
            int paragraphLvl = parentParagraph.Level;

            if (textBoxContainer is Shape { Placeholder: { } } parentShape)
            {
                Placeholder placeholder = (Placeholder)parentShape.Placeholder;
                IFontDataReader phReferencedShape = (IFontDataReader)placeholder.ReferencedShape;
                FontData fontDataPlaceholder = new ();
                if (phReferencedShape != null)
                {
                    phReferencedShape.FillFontData(paragraphLvl, ref fontDataPlaceholder);
                    if (fontDataPlaceholder.FontSize != null)
                    {
                        return fontDataPlaceholder.FontSize / 100;
                    }
                }

                var shapeSlideMaster = parentShape.SlideMasterInternal;

                // From Slide Master body
                if (shapeSlideMaster.TryGetFontSizeFromBody(paragraphLvl, out int fontSizeBody))
                {
                    return fontSizeBody / 100;
                }

                // From Slide Master other
                if (shapeSlideMaster.TryGetFontSizeFromOther(paragraphLvl, out int fontSizeOther))
                {
                    return fontSizeOther / 100;
                }
            }

            // From presentation level
            SCSlideMaster slideMaster = null;
            if (textBoxContainer is Shape shape)
            {
                slideMaster = shape.SlideMasterInternal;
            }
            else
            {
                slideMaster = ((SCTableCell)textBoxContainer).SlideMasterInternal;
            }

            if (slideMaster.Presentation.ParaLvlToFontData.TryGetValue(paragraphLvl, out FontData fontData))
            {
                if (fontData.FontSize != null)
                {
                    return fontData.FontSize / 100;
                }
            }

            return FormatConstants.DefaultFontSize;
        }

        private bool GetBoldFlag()
        {
            A.RunProperties aRunProperties = this.aText.Parent.GetFirstChild<A.RunProperties>();
            if (aRunProperties == null)
            {
                return false;
            }

            if (aRunProperties.Bold != null && aRunProperties.Bold == true)
            {
                return true;
            }

            FontData phFontData = new ();
            FontDataParser.GetFontDataFromPlaceholder(ref phFontData, this.ParentPortion.ParentParagraph);
            if (phFontData.IsBold != null)
            {
                return phFontData.IsBold.Value;
            }

            return false;
        }

        private bool GetItalicFlag()
        {
            A.RunProperties aRunProperties = this.aText.Parent.GetFirstChild<A.RunProperties>();
            if (aRunProperties == null)
            {
                return false;
            }

            if (aRunProperties.Italic != null && aRunProperties.Italic == true)
            {
                return true;
            }

            FontData phFontData = new ();
            FontDataParser.GetFontDataFromPlaceholder(ref phFontData, this.ParentPortion.ParentParagraph);
            if (phFontData.IsItalic != null)
            {
                return phFontData.IsItalic.Value;
            }

            return false;
        }

        private void SetBoldFlag(bool value)
        {
            A.RunProperties aRunPr = this.aText.Parent.GetFirstChild<A.RunProperties>();
            if (aRunPr != null)
            {
                aRunPr.Bold = new BooleanValue(value);
            }
            else
            {
                FontData phFontData = new ();
                FontDataParser.GetFontDataFromPlaceholder(ref phFontData, this.ParentPortion.ParentParagraph);
                if (phFontData.IsBold != null)
                {
                    phFontData.IsBold = new BooleanValue(value);
                }
                else
                {
                    A.EndParagraphRunProperties aEndParaRPr = this.aText.Parent.NextSibling<A.EndParagraphRunProperties>();
                    if (aEndParaRPr != null)
                    {
                        aEndParaRPr.Bold = new BooleanValue(value);
                    }
                    else
                    {
                        aRunPr = new A.RunProperties { Bold = new BooleanValue(value) };
                        this.aText.Parent.InsertAt(aRunPr, 0); // append to <a:r>
                    }
                }
            }
        }

        private void SetItalicFlag(bool isItalic)
        {
            A.RunProperties aRunPr = this.aText.Parent.GetFirstChild<A.RunProperties>();
            if (aRunPr != null)
            {
                aRunPr.Italic = new BooleanValue(isItalic);
            }
            else
            {
                A.EndParagraphRunProperties aEndParaRPr = this.aText.Parent.NextSibling<A.EndParagraphRunProperties>();
                if (aEndParaRPr != null)
                {
                    aEndParaRPr.Italic = new BooleanValue(isItalic);
                }
                else
                {
                    this.aText.Parent.AddRunProperties(isItalic);
                }
            }
        }

        private void SetName(string fontName)
        {
            Shape parentShape = (Shape)this.ParentPortion.ParentParagraph.ParentTextBox.TextFrameContainer;
            if (parentShape.Placeholder != null)
            {
                throw new PlaceholderCannotBeChangedException();
            }

            A.LatinFont latinFont = this.latinFont.Value;
            latinFont.Typeface = fontName;
            this.latinFont.Reset();
        }

        private void SetFontSize(int newFontSize)
        {
            var aRunPr = this.aText.Parent!.GetFirstChild<A.RunProperties>();
            if (aRunPr == null)
            {
                const string errorMsg =
                    "The property value cannot be changed on the Slide level since it belongs to Slide Master. " +
                    "Hence, you should change it on Slide Master level. " +
                    "Note: you can check whether the property can be changed via {property_name}CanBeChanged method.";
                throw new SlideMasterPropertyCannotBeChanged(errorMsg);
            }

            aRunPr.FontSize = newFontSize * 100;
        }
    }
}