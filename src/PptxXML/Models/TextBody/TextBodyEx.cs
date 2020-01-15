using System.Collections.Generic;
using System.Linq;
using ObjectEx.Utilities;
using PptxXML.Models.Settings;
using PptxXML.Services.Builders;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxXML.Models.TextBody
{
    /// <summary>
    /// Represents a text body of the shape.
    /// </summary>
    public class TextBodyEx
    {
        #region Fields

        private readonly IParagraphExBuilder _paragraphExBuilder;
        private readonly ShapeSettings _spSettings;

        #endregion

        #region Properties

        /// <summary>
        /// Gets paragraphs.
        /// </summary>
        public IList<ParagraphEx> Paragraphs { get; private set; }

        #endregion Properties

        #region Constructors

        private TextBodyEx(IParagraphExBuilder paragraphExBuilder, ShapeSettings spSettings, P.TextBody xmlTxtBody)
        {
            Check.NotNull(paragraphExBuilder, nameof(paragraphExBuilder));
            Check.NotNull(spSettings, nameof(spSettings));
            Check.NotNull(spSettings, nameof(xmlTxtBody));
            _paragraphExBuilder = paragraphExBuilder;
            _spSettings = spSettings;
            ParseParagraphs(xmlTxtBody);
        }

        #endregion Constructors

        #region Private Methods

        private void ParseParagraphs(P.TextBody xmlTxtBody)
        {
            var aParagraphs = xmlTxtBody.Descendants<A.Paragraph>();
            Paragraphs = new List<ParagraphEx>(aParagraphs.Count());
            foreach (var p in aParagraphs)
            {
                Paragraphs.Add(_paragraphExBuilder.Build(p, _spSettings));
            }
        }

        #endregion Private Methods

        #region Builder

        public class TextBodyExBuilder : ITextBodyExBuilder
        {
            private readonly IParagraphExBuilder _paragraphExBuilder;

            public TextBodyExBuilder(IParagraphExBuilder paragraphExBuilder)
            {
                Check.NotNull(paragraphExBuilder, nameof(paragraphExBuilder));
                _paragraphExBuilder = paragraphExBuilder;
            }

            public TextBodyEx Build(P.TextBody xmlTxtBody, ShapeSettings spSettings)
            {
                return new TextBodyEx(_paragraphExBuilder, spSettings, xmlTxtBody);
            }
        }

        #endregion
    }
}
