using System.Collections.Generic;
using DocumentFormat.OpenXml;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace SlideXML.Services.Placeholders
{
    /// <summary>
    /// Represents a data of a placeholder.
    /// </summary>
    public class PlaceholderSL
    {
        private Dictionary<int, int> _fontHeights;

        #region Properties

        /// <summary>
        /// Returns placeholder identifier or null.
        /// </summary>
        public int? Id { get; set; } //TODO: refactor: maybe better create two placeholder types: identifier exist and type exist

        /// <summary>
        /// Gets or sets X-coordinate's value.
        /// </summary>
        public long X { get; set; }

        /// <summary>
        /// Gets or sets Y-coordinate's value.
        /// </summary>
        public long Y { get; set; }

        /// <summary>
        /// Gets or sets width value.
        /// </summary>
        public long Width { get; set; }

        /// <summary>
        /// Gets or sets height value.
        /// </summary>
        public long Height { get; set; }

        /// <summary>
        /// Gets or set geometry form code.
        /// </summary>
        public int GeometryCode { get; set; }

        /// <summary>
        /// Gets or sets paragraph level font height.
        /// </summary>
        public Dictionary<int, int> FontHeights
        {
            get
            {
                if (_fontHeights == null)
                {
                    ParseFontHeights();
                }

                return _fontHeights;
            }
        }

        /// <summary>
        /// Gets or sets layout's <see cref="OpenXmlCompositeElement"/> instance.
        /// </summary>
        public OpenXmlCompositeElement CompositeElement { get; set; }

        /// <summary>
        /// Returns placeholder type or null.
        /// </summary>
        public P.PlaceholderValues? Type { get; set; }

        #endregion

        private void ParseFontHeights()
        {
            _fontHeights = new Dictionary<int, int>();
            var shape = (P.Shape)CompositeElement;

            var listStyle = shape.TextBody.ListStyle;
            if (listStyle?.Level1ParagraphProperties != null)
            {
                _fontHeights.Add(1, listStyle.Level1ParagraphProperties.GetFirstChild<A.DefaultRunProperties>().FontSize.Value);

                if (listStyle.Level2ParagraphProperties != null)
                {
                    _fontHeights.Add(2, listStyle.Level2ParagraphProperties.GetFirstChild<A.DefaultRunProperties>().FontSize.Value);
                }
                if (listStyle.Level3ParagraphProperties != null)
                {
                    _fontHeights.Add(3, listStyle.Level3ParagraphProperties.GetFirstChild<A.DefaultRunProperties>().FontSize.Value);
                }
                if (listStyle.Level4ParagraphProperties != null)
                {
                    _fontHeights.Add(4, listStyle.Level4ParagraphProperties.GetFirstChild<A.DefaultRunProperties>().FontSize.Value);
                }
                if (listStyle.Level5ParagraphProperties != null)
                {
                    _fontHeights.Add(5, listStyle.Level5ParagraphProperties.GetFirstChild<A.DefaultRunProperties>().FontSize.Value);
                }
                if (listStyle.Level6ParagraphProperties != null)
                {
                    _fontHeights.Add(6, listStyle.Level6ParagraphProperties.GetFirstChild<A.DefaultRunProperties>().FontSize.Value);
                }
                if (listStyle.Level7ParagraphProperties != null)
                {
                    _fontHeights.Add(7, listStyle.Level7ParagraphProperties.GetFirstChild<A.DefaultRunProperties>().FontSize.Value);
                }
                if (listStyle.Level8ParagraphProperties != null)
                {
                    _fontHeights.Add(8, listStyle.Level8ParagraphProperties.GetFirstChild<A.DefaultRunProperties>().FontSize.Value);
                }
                if (listStyle.Level9ParagraphProperties != null)
                {
                    _fontHeights.Add(9, listStyle.Level9ParagraphProperties.GetFirstChild<A.DefaultRunProperties>().FontSize.Value);
                }
            }
            else
            {
                _fontHeights.Add(1, shape.TextBody.GetFirstChild<A.Paragraph>().GetFirstChild<A.EndParagraphRunProperties>().FontSize.Value);
            }
        }
    }
}
