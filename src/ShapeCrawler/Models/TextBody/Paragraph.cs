using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Extensions;
using ShapeCrawler.Models.Settings;
using ShapeCrawler.Statics;
using A = DocumentFormat.OpenXml.Drawing;
// ReSharper disable PossibleMultipleEnumeration

namespace ShapeCrawler.Models.TextBody
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
        private readonly Lazy<Bullet> _bullet;

        #endregion Fields

        #region Properties

        /// <summary>
        /// Gets paragraph's text string.
        /// </summary>
        public string Text => _text.Value;

        /// <summary>
        /// Gets paragraph text portions.
        /// </summary>
        public IList<Portion> Portions => _portions.Value;

        /// <summary>
        /// Gets paragraph bullet. Returns null if bullet does not exist.
        /// </summary>
        public Bullet Bullet => _bullet.Value;

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
            _bullet = new Lazy<Bullet>(GetBullet);
        }

        private Bullet GetBullet()
        {
            return new Bullet(_xmlParagraph.ParagraphProperties);
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

        private List<Portion> GetPortions()
        {
            var runs = _xmlParagraph.Elements<A.Run>();
            var resultPortions = runs.Any() ? PortionsFromRuns(runs) : PortionsFromField();

            return resultPortions;
        }

        private List<Portion> PortionsFromRuns(IEnumerable<A.Run> runs)
        {
            var portions = new List<Portion>(runs.Count());
            foreach (var run in runs)
            {
                var fh = FontHeightFromRun(run);
                portions.Add(new Portion(run.Text.Text, fh));
            }
            return portions;
        }

        private List<Portion> PortionsFromField()
        {
            var text = _xmlParagraph.GetFirstChild<A.Field>().GetFirstChild<A.Text>().Text;
            var fh = FontHeightFromOther();
            var portions = new List<Portion>(1)
            {
                new Portion(text, fh)
            };

            return portions;
        }

        private int FontHeightFromRun(A.Run run)
        {
            var runFs = run.RunProperties?.FontSize;
            var resultFh = runFs == null ? FontHeightFromOther() : runFs.Value;

            return resultFh;
        }

        private int FontHeightFromOther()
        {
            // if element is placeholder, tries to get from placeholder data
            var xmlElement = _spContext.SdkElement;
            if (xmlElement.IsPlaceholder())
            {
                var prFontHeight = _spContext.PlaceholderFontService.TryGetFontHeight((OpenXmlCompositeElement)xmlElement, _innerPrLvl.Value);
                if (prFontHeight != null)
                {
                    return (int)prFontHeight;
                }
            }

            if (_spContext.PreSettings.LlvFontHeights.ContainsKey(_innerPrLvl.Value))
            {
                return _spContext.PreSettings.LlvFontHeights[_innerPrLvl.Value];
            }

            var exist = _spContext.TryGetFontHeight(_innerPrLvl.Value, out int fh);
            if (exist)
            {
                return fh;
            }

            return FormatConstants.DefaultFontHeight;
        }

        #endregion Private Methods
    }
}
