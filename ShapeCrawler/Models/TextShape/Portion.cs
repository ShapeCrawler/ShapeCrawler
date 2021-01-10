using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml;
using ShapeCrawler.Extensions;
using ShapeCrawler.Settings;
using ShapeCrawler.Shared;
using ShapeCrawler.Statics;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Models.TextShape
{
    /// <summary>
    /// Represents a text paragraph portion.
    /// </summary>
    [SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
    public class Portion
    {
        private readonly ResettableLazy<FontSc> _font;
        private readonly ShapeContext _shapeContext;
        private readonly int _innerPrLvl;

        #region Internal Properties

        internal ParagraphEx ParagraphEx { get; }
        internal readonly A.Text AText;

        #endregion Internal Properties

        #region Public Properties

        /// <summary>
        /// Gets or sets text.
        /// </summary>
        public string Text
        {
            get => AText.Text;
            set => AText.Text = value;
        }

        /// <summary>
        /// Gets font.
        /// </summary>
        public FontSc Font => _font.Value;
        
        /// <summary>
        /// Removes portion from the paragraph.
        /// </summary>
        public void Remove()
        {
            ParagraphEx.Portions.Remove(this);
        }

        #endregion Public Properties

        #region Constructors

        public Portion(A.Text aText, ParagraphEx paragraphEx, ShapeContext shapeContext, int innerPrLvl)
        {
            AText = aText;
            ParagraphEx = paragraphEx;
            _shapeContext = shapeContext;
            _innerPrLvl = innerPrLvl;
            _font = new ResettableLazy<FontSc>(CreateFont);
        }

        #endregion Constructors

        private FontSc CreateFont()
        {
            Int32Value sdkFontSize = AText.Parent.GetFirstChild<A.RunProperties>()?.FontSize;
            int fontSize = sdkFontSize != null ? sdkFontSize.Value : FontSizeFromOther();

            return new FontSc(AText, fontSize, this);
        }

        private int FontSizeFromOther()
        {
            // If element is placeholder, tries to get from placeholder data
            OpenXmlElement sdkElement = _shapeContext.SdkElement;
            if (sdkElement.IsPlaceholder())
            {
                var prFontHeight =
                    _shapeContext.PlaceholderFontService.TryGetFontHeight((OpenXmlCompositeElement)sdkElement,
                        _innerPrLvl);
                if (prFontHeight != null)
                {
                    return (int)prFontHeight;
                }
            }

            if (_shapeContext.presentationData.LlvFontHeights.ContainsKey(_innerPrLvl))
            {
                return _shapeContext.presentationData.LlvFontHeights[_innerPrLvl];
            }

            var exist = _shapeContext.TryGetFontSize(_innerPrLvl, out int fh);
            if (exist)
            {
                return fh;
            }

            return FormatConstants.DefaultFontSize;
        }
    }
}