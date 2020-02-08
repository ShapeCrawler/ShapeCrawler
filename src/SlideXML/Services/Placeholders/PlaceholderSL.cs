using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SlideXML.Enums;
using SlideXML.Models.TextBody;
using SlideXML.Validation;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace SlideXML.Services.Placeholders
{
    /// <summary>
    /// Represents placeholder data on layout/master slide.
    /// </summary>
    public class PlaceholderSL : IEquatable<PlaceholderSL>
    {
        private Dictionary<int, int> _fontHeights; //TODO: set lazy

        #region Properties

        /// <summary>
        /// Returns placeholder type.
        /// </summary>
        public PlaceholderType Type { get; }

        /// <summary>
        /// Returns placeholder index for custom type; Null will be returned for pre-define placeholder types.
        /// </summary>
        public int? Index { get; } //TODO: consider to split on Predefine and Custom placeholders

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
        /// Gets or sets <see cref="OpenXmlCompositeElement"/> instance of layout/master.
        /// </summary>
        public OpenXmlCompositeElement CompositeElement { get; set; }

        /// <summary>
        /// Gets or sets the <see cref="SlideLayoutPart"/> instance.
        /// </summary>
        public SlideLayoutPart SlideLayoutPart { get; set; }

        #endregion

        #region Constructors

        /// <summary>
        /// Creates a new <see cref="PlaceholderSL"/> instance from <see cref="PlaceholderXML"/>.
        /// </summary>
        public PlaceholderSL(PlaceholderXML phXml)
        {
            Check.NotNull(phXml, nameof(phXml));
            Type = phXml.PlaceholderType;
            Index = phXml.Index;
        }

        #endregion Constructors

        #region Public Methods

        /// <summary>
        /// Indicates whether the current object is equal to another <see cref="PlaceholderSL"/> instance.
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        [SuppressMessage("ReSharper", "PossibleNullReferenceException")]
        public bool Equals(PlaceholderSL other)
        {
            Check.NotNull(other, nameof(other));

            // Compares custom type
            if (Type == PlaceholderType.Custom && other.Type == PlaceholderType.Custom && Index == other.Index)
            {
                return true;
            }

            // Compares non-custom type
            return Type != PlaceholderType.Custom && Type == other.Type;
        }

        /// <summary>
        /// Indicates whether the current object is equal to another <see cref="Object"/> instance.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            Check.NotNull(obj, nameof(obj));
            var ph = (PlaceholderSL)obj;

            return Equals(ph);
        }

        /// <summary>
        /// Returns the hash calculating upon the formula suggested here: https://stackoverflow.com/a/263416/2948684
        /// </summary>
        /// <remarks></remarks>
        public override int GetHashCode()
        {
            var hash = 17;

            // For pre-define type
            if (Type != PlaceholderType.Custom)
            {
                hash = hash + 23 + Type.GetHashCode();
            }
            else
            {
                // For custom type
                hash = hash + 23 + Type.GetHashCode();
                hash = hash + 23 + Index.GetHashCode();
            }

            return hash;
        }

        #endregion

        #region Private Methods

        private void ParseFontHeights()
        {
            _fontHeights = new Dictionary<int, int>();
            if (Type.Equals(PlaceholderType.Title)) // for title placeholder font height is parsed from slide master
            {
                _fontHeights.Add(1, SlideLayoutPart.SlideMasterPart.SlideMaster.TextStyles.TitleStyle.Level1ParagraphProperties.GetFirstChild<A.DefaultRunProperties>().FontSize.Value);
            }
            else // non-title placeholder
            {
                var shape = (P.Shape)CompositeElement;
                foreach (var textPr in shape.TextBody.Descendants<A.TextParagraphPropertiesType>())
                {
                    var fs = textPr.GetFirstChild<A.DefaultRunProperties>()?.FontSize;
                    if (fs == null)
                    {
                        continue;
                    }
                    // fourth character of LocalName contains level number, example: "lvl1pPr, lvl2pPr, etc."
                    var lvl = int.Parse(textPr.LocalName[3].ToString()); 
                    _fontHeights.Add(lvl, fs.Value);
                }
                if (!_fontHeights.Any())
                {
                    _fontHeights.Add(1, shape.TextBody.GetFirstChild<A.Paragraph>().GetFirstChild<A.EndParagraphRunProperties>().FontSize.Value);
                }
            }
        }

        #endregion
    }
}
