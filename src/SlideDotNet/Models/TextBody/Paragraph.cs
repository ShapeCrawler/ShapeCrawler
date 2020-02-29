using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using SlideDotNet.Extensions;
using SlideDotNet.Models.Settings;
using SlideDotNet.Statics;
using A = DocumentFormat.OpenXml.Drawing;
// ReSharper disable PossibleMultipleEnumeration

namespace SlideDotNet.Models.TextBody
{
    /// <summary>
    /// Represents a text paragraph.
    /// </summary>
    public class Paragraph
    {
        #region Fields

        private readonly IShapeContext _spContext;

        private readonly A.Paragraph _xmlParagraph;
        private readonly Lazy<int> _innerPrLvl; // inner paragraph level started from one
        private readonly Lazy<string> _text;
        private readonly Lazy<List<Portion>> _portions;

        #endregion Fields

        #region Properties

        /// <summary>
        /// Returns the paragraph's text string.
        /// </summary>
        public string Text => _text.Value;

        /// <summary>
        /// Returns paragraph text portions.
        /// </summary>
        public IList<Portion> Portions => _portions.Value;

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initializes an instance of the <see cref="Paragraph"/> class.
        /// </summary>
        /// <param name="spContext"></param>
        /// <param name="xmlParagraph">A XML paragraph which contains a text.</param>
        public Paragraph(IShapeContext spContext, A.Paragraph xmlParagraph)
        {
            _xmlParagraph = xmlParagraph ?? throw new ArgumentNullException(nameof(xmlParagraph));
            _spContext = spContext ?? throw new ArgumentNullException(nameof(spContext));
            _innerPrLvl = new Lazy<int>(GetInnerLevel(_xmlParagraph));
            _text = new Lazy<string>(GetText);
            _portions = new Lazy<List<Portion>>(GetPortions);
        }

        #endregion Constructors

        #region Private Methods

        private static int GetInnerLevel(A.Paragraph xmlParagraph)
        {
            // XML-paragraph enumeration started from zero. Null is also zero
            var outerLvl = xmlParagraph.ParagraphProperties?.Level ?? 0;
            var interLvl = outerLvl + 1;

            return interLvl;
        }

        private string GetText()
        {
            return Portions.Select(p => p.Text).Aggregate((result, next) => result + next);
        }

        [SuppressMessage("ReSharper", "ConvertIfStatementToConditionalTernaryExpression")]
        private List<Portion> GetPortions()
        {
            var runs = _xmlParagraph.Elements<A.Run>();
            if (runs.Any())
            {
                var portions = new List<Portion>(runs.Count());
                foreach (var run in runs)
                {
                    var runFh = GetRunFontHeight(run);
                    portions.Add(new Portion(run.Text.Text, runFh));
                }
                return portions;
            }
            else
            {
                var text = _xmlParagraph.Descendants<A.Text>().Single().Text; // text container candidate is <a:fld> element
                var portions = new List<Portion>(1)
                {
                    new Portion(text)
                };
                return portions;
            }
        }

        private int GetRunFontHeight(A.Run run)
        {
            // first, tries parse font height from current run (portion)
            var runFs = run.RunProperties?.FontSize;
            if (runFs != null)
            {
                return runFs.Value;
            }

            // if element is placeholder, tries to get from placeholder data
            var xmlElement = _spContext.XmlElement;
            if (xmlElement.IsPlaceholder())
            {
                var prFontHeight = _spContext.PlaceholderFontService.TryGetHeight(xmlElement, _innerPrLvl.Value);
                if (prFontHeight != null)
                {
                    return (int)prFontHeight;
                }
            }

            if (_spContext.PreSettings.LlvFontHeights.ContainsKey(_innerPrLvl.Value))
            {
                return _spContext.PreSettings.LlvFontHeights[_innerPrLvl.Value];
            }

            var exist = _spContext.TryFromMasterOther(_innerPrLvl.Value, out int fh);
            if (exist)
            {
                return fh;
            }

            return FormatConstants.DefaultFontHeight;
        }

        #endregion Private Methods
    }
}
