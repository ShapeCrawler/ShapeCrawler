using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using ShapeCrawler.Settings;

namespace ShapeCrawler.Texts
{
    // TODO: Override ToString()
    [SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
    public sealed class TextSc
    {
        #region Fields

        private readonly ShapeContext _spContext;
        private readonly Lazy<string> _content;
        private readonly OpenXmlCompositeElement _compositeElement;

        #endregion Fields

        internal ShapeSc ShapeEx { get; }

        #region Public Properties

        public ParagraphCollection Paragraphs => ParagraphCollection.Parse(_compositeElement, _spContext, this); // TODO: make lazy

        /// <summary>
        /// Gets or sets text string content.
        /// </summary>
        public string Content
        {
            get => _content.Value;
            set => SetContent(value);
        }

        #endregion Public Properties

        #region Constructors

        /// <summary>
        /// Initializes an instance of the <see cref="TextSc"/>.
        /// </summary>
        public TextSc(OpenXmlCompositeElement compositeElement, ShapeSc shapeEx)
            :this(shapeEx.Context, compositeElement)
        {
            ShapeEx = shapeEx;
        }

        public TextSc(ShapeContext spContext, OpenXmlCompositeElement compositeElement)
        {
            _spContext = spContext;
            _compositeElement = compositeElement;
            _content = new Lazy<string>(GetText);
        }

        #endregion Constructors

        #region Private Methods

        private void SetContent(string value)
        {
            if (Paragraphs.Count > 1)
            {
                // Remove all except first paragraph
                Paragraphs.RemoveRange(1, Paragraphs.Count - 1);
            }

            Paragraphs.Single().Text = value;
        }

        private string GetText()
        {
            var sb = new StringBuilder();
            sb.Append(Paragraphs[0].Text);
            
            // If the number of paragraphs more than one
            var numPr = Paragraphs.Count;
            var index = 1;
            while (index < numPr)
            {
                sb.AppendLine();
                sb.Append(Paragraphs[index].Text);

                index++;
            }

            return sb.ToString();
        }

        #endregion Private Methods
    }
}
