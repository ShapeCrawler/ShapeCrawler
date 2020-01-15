using System.Collections.Generic;
using System.Linq;
using ObjectEx.Utilities;
using PptxXML.Models.Settings;
using PptxXML.Services.Builders;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxXML.Models.TextBody
{
    /// <summary>
    /// Represents a text paragraph.
    /// </summary>
    public class ParagraphEx
    {
        #region Fields

        private readonly A.Paragraph _aParagraph;
        private readonly ShapeSettings _shapeSetting;

        private string _text;
        private int? _lvl; // paragraph's level
        private List<Portion> _portions;

        #endregion Fields

        #region Properties

        /// <summary>
        /// Returns the paragraph's text string.
        /// </summary>
        public string Text {
            get
            {
                if (_text == null)
                {
                    InitText();
                }

                return _text;
            }
        } 

        /// <summary>
        /// Returns paragraph text portions.
        /// </summary>
        public IList<Portion> Portions {
            get
            {
                if (_portions == null)
                {
                    InitPortions();
                }

                return _portions;
            }
        }

        #endregion Properties

        #region Constructors

        public ParagraphEx(ShapeSettings shapeSetting, A.Paragraph aParagraph)
        {
            Check.NotNull(aParagraph, nameof(aParagraph));
            Check.NotNull(shapeSetting, nameof(shapeSetting));
            _aParagraph = aParagraph;
            _shapeSetting = shapeSetting;
        }

        #endregion Constructors

        #region Private Methods

        private void InitText()
        {
            _text = _aParagraph.Descendants<A.Text>().Select(t => t.Text).Aggregate((t1, t2) => t1 + t2);
        }

        private void InitPortions()
        {
            var prLvl = ParseLevel();
            var runs = _aParagraph.Elements<A.Run>();
            _portions = new List<Portion>(runs.Count());
            var ph = _shapeSetting.Placeholder;

            if (ph != null) // is placeholder
            {
                var fh = ph.FontHeights[prLvl]; // gets font height from placeholder
                foreach (var run in runs) //TODO: delete unnecessary run
                {
                    _portions.Add(new Portion(fh));
                }
            }
            else // is not placeholder
            {
                foreach (var run in runs)
                {
                    var fh = run.RunProperties?.FontSize?.Value ?? _shapeSetting.PreSettings.LlvFontHeights[prLvl];
                    _portions.Add(new Portion(fh));
                }
            }
        }

        private int ParseLevel()
        {
            if (_lvl == null)
            {
                // Gets default paragraph level font height for current paragraph's level
                _lvl = _aParagraph.ParagraphProperties?.Level?.Value;
                if (_lvl == null)
                {
                    _lvl = 1;
                }
                else
                {
                    // By unknown reason, slide and presentation's default settings have different numbering
                    _lvl++;
                }
            }

            return (int)_lvl;
        }

        #endregion Private Methods

        #region Builder

        public class ParagraphExBuilder : IParagraphExBuilder
        {
            #region Dependencies

            private readonly ShapeSettings _spSettings;

            #endregion Dependencies

            #region Constructors

            public ParagraphExBuilder()
            {

            }

            #endregion Constructors

            public ParagraphEx Build(A.Paragraph aParagraph, ShapeSettings spSetting)
            {
                Check.NotNull(aParagraph, nameof(aParagraph));
                Check.NotNull(spSetting, nameof(spSetting));
                return new ParagraphEx(spSetting, aParagraph);
            }
        }

        #endregion Builder
    }
}
