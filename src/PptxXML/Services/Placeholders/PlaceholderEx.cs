using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxXML.Services.Placeholders
{
    /// <summary>
    /// Represents a data of a placeholder.
    /// </summary>
    public class PlaceholderEx
    {
        private Dictionary<int, int> _fontHeights;

        #region Properties

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
        /// Gets or sets element's <see cref="OpenXmlCompositeElement"/> instance.
        /// </summary>
        public OpenXmlCompositeElement CompositeElement { get; set; }

        #endregion

        private void ParseFontHeights()
        {
            _fontHeights = new Dictionary<int, int>();
            var shape = (P.Shape)CompositeElement;

            // Defines placeholder type
            var ph = CompositeElement.Descendants<P.PlaceholderShape>().FirstOrDefault();
            var phType = ph.Type;

            // Title placeholder type
            if (phType == P.PlaceholderValues.Title || phType == P.PlaceholderValues.CenteredTitle || phType == P.PlaceholderValues.SubTitle)
            {
                _fontHeights.Add(1, shape.TextBody.GetFirstChild<A.Paragraph>().GetFirstChild<A.EndParagraphRunProperties>().FontSize.Value);
            }
            else // Other placeholder types
            {
                var listStyle = shape.TextBody.ListStyle;
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
        }
    }
}
