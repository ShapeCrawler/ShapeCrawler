using System;
using DocumentFormat.OpenXml;
using ShapeCrawler.Drawing;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Extensions;
using ShapeCrawler.Factories;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Shared;
using ShapeCrawler.Statics;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.AutoShapes
{
    internal class SCFont : IFont
    {
        internal readonly Portion Portion;

        private readonly A.Text aText;
        private readonly A.FontScheme aFontScheme;
        private readonly Lazy<ColorFormat> colorFormat;
        private readonly ResettableLazy<A.LatinFont> latinFont;
        private readonly ResettableLazy<int> size;

        internal SCFont(A.Text aText, Portion portion)
        {
            this.aText = aText;
            this.size = new ResettableLazy<int>(this.GetSize);
            this.latinFont = new ResettableLazy<A.LatinFont>(this.GetALatinFont);
            this.colorFormat = new Lazy<ColorFormat>(() => new ColorFormat(this));
            this.Portion = portion;
            this.aFontScheme = ((Shape)portion.ParentParagraph.ParentTextBox.ParentTextBoxContainer).ThemePart.Theme.ThemeElements.FontScheme;
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

        public IColorFormat ColorFormat => this.colorFormat.Value;

        public bool SizeCanBeChanged()
        {
            A.RunProperties runPr = this.aText.Parent.GetFirstChild<A.RunProperties>();
            return runPr != null;
        }

        #endregion Public Properties

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
            FontDataParser.GetFontDataFromPlaceholder(ref phFontData, this.Portion.ParentParagraph);
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
            Int32Value aRunPrFontSize = this.Portion.AText.Parent.GetFirstChild<A.RunProperties>()?.FontSize;
            if (aRunPrFontSize != null)
            {
                return aRunPrFontSize.Value;
            }

            Shape fontParentShape = (Shape)this.Portion.ParentParagraph.ParentTextBox.ParentTextBoxContainer;
            int paragraphLvl = this.Portion.ParentParagraph.Level;

            // Try get font size from placeholder
            if (fontParentShape.Placeholder != null)
            {
                Placeholder placeholder = (Placeholder)fontParentShape.Placeholder;
                IFontDataReader phReferencedShape = (IFontDataReader)placeholder.ReferencedShape;
                FontData fontDataPlaceholder = new ();
                if (phReferencedShape != null)
                {
                    phReferencedShape.FillFontData(paragraphLvl, ref fontDataPlaceholder);
                    if (fontDataPlaceholder.FontSize != null)
                    {
                        return fontDataPlaceholder.FontSize;
                    }
                }

                // From Slide Master body
                if (fontParentShape.SlideMaster.TryGetFontSizeFromBody(paragraphLvl, out int fontSizeBody))
                {
                    return fontSizeBody;
                }

                // From Slide Master other
                if (fontParentShape.SlideMaster.TryGetFontSizeFromOther(paragraphLvl, out int fontSizeOther))
                {
                    return fontSizeOther;
                }
            }

            // From presentation level
            if (fontParentShape.ParentPresentation.ParaLvlToFontData.TryGetValue(paragraphLvl, out FontData fontData))
            {
                if (fontData.FontSize != null)
                {
                    return fontData.FontSize;
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

            FontData phFontData = new();
            FontDataParser.GetFontDataFromPlaceholder(ref phFontData, this.Portion.ParentParagraph);
            if (phFontData.IsBold != null)
            {
                return phFontData.IsBold.Value;
            }

            return false;
        }

        private bool GetItalicFlag()
        {
            A.RunProperties aRunProperties = aText.Parent.GetFirstChild<A.RunProperties>();
            if (aRunProperties == null)
            {
                return false;
            }

            if (aRunProperties.Italic != null && aRunProperties.Italic == true)
            {
                return true;
            }

            FontData phFontData = new();
            FontDataParser.GetFontDataFromPlaceholder(ref phFontData, Portion.ParentParagraph);
            if (phFontData.IsItalic != null)
            {
                return phFontData.IsItalic.Value;
            }

            return false;
        }

        private void SetBoldFlag(bool value)
        {
            A.RunProperties aRunPr = aText.Parent.GetFirstChild<A.RunProperties>();
            if (aRunPr != null)
            {
                aRunPr.Bold = new BooleanValue(value);
            }
            else
            {
                FontData phFontData = new();
                FontDataParser.GetFontDataFromPlaceholder(ref phFontData, Portion.ParentParagraph);
                if (phFontData.IsBold != null)
                {
                    phFontData.IsBold = new BooleanValue(value);
                }
                else
                {
                    A.EndParagraphRunProperties aEndParaRPr = aText.Parent.NextSibling<A.EndParagraphRunProperties>();
                    if (aEndParaRPr != null)
                    {
                        aEndParaRPr.Bold = new BooleanValue(value);
                    }
                    else
                    {
                        aRunPr = new A.RunProperties {Bold = new BooleanValue(value)};
                        aText.Parent.InsertAt(aRunPr, 0); // append to <a:r>
                    }
                }
            }
        }

        private void SetItalicFlag(bool value)
        {
            A.RunProperties aRunPr = aText.Parent.GetFirstChild<A.RunProperties>();
            if (aRunPr != null)
            {
                aRunPr.Italic = new BooleanValue(value);
            }
            else
            {
                A.EndParagraphRunProperties aEndParaRPr = aText.Parent.NextSibling<A.EndParagraphRunProperties>();
                if (aEndParaRPr != null)
                {
                    aEndParaRPr.Italic = new BooleanValue(value);
                }
                else
                {
                    aRunPr = new A.RunProperties {Italic = new BooleanValue(value)};
                    this.aText.Parent.InsertAt(aRunPr, 0); // append to <a:r>
                }
            }
        }

        private void SetName(string fontName)
        {
            Shape parentShape = (Shape)this.Portion.ParentParagraph.ParentTextBox.ParentTextBoxContainer;
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
            A.RunProperties aRunPr = this.aText.Parent.GetFirstChild<A.RunProperties>();
            if (aRunPr == null)
            {
                const string errorMsg =
                    "The property value cannot be changed on the Slide level since it belongs to Slide Master. " +
                    "Hence, you should change it on Slide Master level. " +
                    "Note: you can check whether the property can be changed via {property_name}CanBeChanged method.";
                throw new SlideMasterPropertyCannotBeChanged(errorMsg);
            }

            aRunPr.FontSize = newFontSize;
        }

        private void SetSolidColorHex(string value)
        {
            A.RunProperties aRunPr = this.aText.Parent.GetFirstChild<A.RunProperties>();
            if (aRunPr != null)
            {
                var aSolidFill = new A.SolidFill
                {
                    RgbColorModelHex = new A.RgbColorModelHex {Val = value}
                };

                aRunPr.SolidFill()?.Remove(); // remove old color
                aRunPr.InsertAt(aSolidFill, 0);
            }
            else
            {
                var aSolidFill = new A.SolidFill
                {
                    RgbColorModelHex = new A.RgbColorModelHex
                    {
                        Val = value
                    }
                };

                aRunPr = new A.RunProperties(aSolidFill);
                this.aText.Parent.InsertAt(aRunPr, 0);
            }
        }
    }
}