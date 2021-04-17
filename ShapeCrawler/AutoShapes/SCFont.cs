using System;
using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml;
using ShapeCrawler.Drawing;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Extensions;
using ShapeCrawler.Factories;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Settings;
using ShapeCrawler.Shared;
using ShapeCrawler.Statics;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.AutoShapes
{
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    internal class SCFont : IFont
    {
        private readonly A.Text _aText;
        private readonly Lazy<ColorFormat> _colorFormat;
        private readonly ResettableLazy<A.LatinFont> _latinFont;
        private readonly int _paragraphLvl;
        private readonly ResettableLazy<int> _size;
        internal readonly Portion Portion;

        #region Constructors

        internal SCFont(A.Text aText, Portion portion)
        {
            _aText = aText;
            _size = new ResettableLazy<int>(GetSize);
            _latinFont = new ResettableLazy<A.LatinFont>(GetALatinFont);
            _colorFormat = new Lazy<ColorFormat>(() => new ColorFormat(this));
            _paragraphLvl = portion.Paragraph.Level;
            Portion = portion;
        }

        #endregion Constructors

        #region Public Properties

        /// <summary>
        ///     Gets font name.
        /// </summary>
        public string Name
        {
            get => GetName();
            set => SetName(value);
        }

        /// <summary>
        ///     Gets or sets font size in EMUs.
        /// </summary>
        public int Size
        {
            get => _size.Value;
            set => SetFontSize(value);
        }

        public bool IsBold
        {
            get => GetBoldFlag();
            set => SetBoldFlag(value);
        }

        public bool IsItalic
        {
            get => GetItalicFlag();
            set => SetItalicFlag(value);
        }

        public IColorFormat ColorFormat => _colorFormat.Value;

        /// <summary>
        ///     Gets value indicating whether font size can be changed.
        /// </summary>
        public bool SizeCanBeChanged()
        {
            var runPr = _aText.Parent.GetFirstChild<A.RunProperties>();
            return runPr != null;
        }

        #endregion Public Properties

        #region Private Methods

        private string GetName()
        {
            const string majorLatinFont = "+mj-lt";
            if (_latinFont.Value.Typeface == majorLatinFont)
            {
                return Portion.Paragraph.TextBox.AutoShape.ThemePart.Theme.ThemeElements.FontScheme.MajorFont.LatinFont
                    .Typeface;
            }

            return _latinFont.Value.Typeface;
        }

        private A.LatinFont GetALatinFont()
        {
            A.RunProperties aRunProperties = _aText.Parent.GetFirstChild<A.RunProperties>();
            A.LatinFont aLatinFont = aRunProperties?.GetFirstChild<A.LatinFont>();

            if (aLatinFont != null)
            {
                return aLatinFont;
            }

            FontData phFontData = new();
            FontDataParser.GetFontDataFromPlaceholder(ref phFontData, Portion.Paragraph);
            {
                if (phFontData.ALatinFont != null)
                {
                    return phFontData.ALatinFont;
                }
            }

            // Get from theme
            return Portion.Paragraph.TextBox.AutoShape.ThemePart.Theme.ThemeElements.FontScheme.MinorFont.LatinFont;
        }

        private int GetSize()
        {
            Int32Value aRunPrFontSize = Portion.AText.Parent.GetFirstChild<A.RunProperties>()?.FontSize;
            if (aRunPrFontSize != null)
            {
                return aRunPrFontSize.Value;
            }

            Shape fontParentShape = Portion.Paragraph.TextBox.AutoShape;
            int paragraphLvl = Portion.Paragraph.Level;

            // Try get font size from placeholder
            if (fontParentShape.Placeholder != null)
            {
                Placeholder placeholder = (Placeholder) fontParentShape.Placeholder;
                IFontDataReader phReferencedShape = (IFontDataReader) placeholder.ReferencedShape;
                FontData fontDataPlaceholder = new();
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
            if (fontParentShape.Presentation.ParaLvlToFontData.TryGetValue(paragraphLvl, out FontData fontData))
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
            A.RunProperties aRunProperties = _aText.Parent.GetFirstChild<A.RunProperties>();
            if (aRunProperties == null)
            {
                return false;
            }

            if (aRunProperties.Bold != null && aRunProperties.Bold == true)
            {
                return true;
            }

            FontData phFontData = new();
            FontDataParser.GetFontDataFromPlaceholder(ref phFontData, Portion.Paragraph);
            if (phFontData.IsBold != null)
            {
                return phFontData.IsBold.Value;
            }

            return false;
        }

        private bool GetItalicFlag()
        {
            A.RunProperties aRunProperties = _aText.Parent.GetFirstChild<A.RunProperties>();
            if (aRunProperties == null)
            {
                return false;
            }

            if (aRunProperties.Italic != null && aRunProperties.Italic == true)
            {
                return true;
            }

            FontData phFontData = new();
            FontDataParser.GetFontDataFromPlaceholder(ref phFontData, Portion.Paragraph);
            if (phFontData.IsItalic != null)
            {
                return phFontData.IsItalic.Value;
            }

            return false;
        }

        private void SetBoldFlag(bool value)
        {
            A.RunProperties aRunPr = _aText.Parent.GetFirstChild<A.RunProperties>();
            if (aRunPr != null)
            {
                aRunPr.Bold = new BooleanValue(value);
            }
            else
            {
                FontData phFontData = new();
                FontDataParser.GetFontDataFromPlaceholder(ref phFontData, Portion.Paragraph);
                if (phFontData.IsBold != null)
                {
                    phFontData.IsBold = new BooleanValue(value);
                }
                else
                {
                    A.EndParagraphRunProperties aEndParaRPr = _aText.Parent.NextSibling<A.EndParagraphRunProperties>();
                    if (aEndParaRPr != null)
                    {
                        aEndParaRPr.Bold = new BooleanValue(value);
                    }
                    else
                    {
                        aRunPr = new A.RunProperties {Bold = new BooleanValue(value)};
                        _aText.Parent.InsertAt(aRunPr, 0); // append to <a:r>
                    }
                }
            }
        }

        private void SetItalicFlag(bool value)
        {
            A.RunProperties aRunPr = _aText.Parent.GetFirstChild<A.RunProperties>();
            if (aRunPr != null)
            {
                aRunPr.Italic = new BooleanValue(value);
            }
            else
            {
                A.EndParagraphRunProperties aEndParaRPr = _aText.Parent.NextSibling<A.EndParagraphRunProperties>();
                if (aEndParaRPr != null)
                {
                    aEndParaRPr.Italic = new BooleanValue(value);
                }
                else
                {
                    aRunPr = new A.RunProperties {Italic = new BooleanValue(value)};
                    _aText.Parent.InsertAt(aRunPr, 0); // append to <a:r>
                }
            }
        }

        private void SetName(string fontName)
        {
            if (Portion.Paragraph.TextBox.AutoShape.Placeholder != null)
            {
                throw new PlaceholderCannotBeChangedException();
            }

            A.LatinFont latinFont = _latinFont.Value;
            latinFont.Typeface = fontName;
            _latinFont.Reset();
        }

        private void SetFontSize(int newFontSize)
        {
            A.RunProperties aRunPr = _aText.Parent.GetFirstChild<A.RunProperties>();
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
            A.RunProperties aRunPr = _aText.Parent.GetFirstChild<A.RunProperties>();
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
                        Val = value,
                    },
                };

                aRunPr = new A.RunProperties(aSolidFill);
                _aText.Parent.InsertAt(aRunPr, 0);
            }
        }

        #endregion Private Methods
    }
}