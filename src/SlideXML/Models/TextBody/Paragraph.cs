using System;
using System.Collections.Generic;
using System.Linq;
using SlideXML.Models.Settings;
using SlideXML.Validation;
using A = DocumentFormat.OpenXml.Drawing;
// ReSharper disable PossibleMultipleEnumeration

namespace SlideXML.Models.TextBody
{
    /// <summary>
    /// Represents a text paragraph.
    /// </summary>
    public class Paragraph
    {
        #region Fields

        private readonly A.Paragraph _aParagraph;
        private readonly ElementSettings _shapeSetting;
        private string _text;
        private readonly Lazy<int> _lvl;
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

        /// <summary>
        /// Initializes an instance of the <see cref="Paragraph"/> class.
        /// </summary>
        /// <param name="elSetting"></param>
        /// <param name="aParagraph">A XML paragraph which contains a text.</param>
        public Paragraph(ElementSettings elSetting, A.Paragraph aParagraph)
        {
            Check.NotNull(aParagraph, nameof(aParagraph));
            Check.NotNull(elSetting, nameof(elSetting));
            _aParagraph = aParagraph;
            _shapeSetting = elSetting;
            _lvl = new Lazy<int>(ParseLevel);
        }

        #endregion Constructors

        #region Private Methods

        private void InitText()
        {
            _text = Portions.Select(p => p.Text).Aggregate((t1, t2) => t1 + t2);
        }

        private void InitPortions()
        {
            var runs = _aParagraph.Elements<A.Run>();
            _portions = new List<Portion>(runs.Count());
            var ph = _shapeSetting.Placeholder;

            foreach (var run in runs)
            {
                // First tries to get font height from run, then placeholder and only then from presentation settings.
                var fh = run.RunProperties?.FontSize?.Value ?? ph?.FontHeights[_lvl.Value] ?? _shapeSetting.PreSettings.LlvFontHeights[_lvl.Value];
                
                _portions.Add(new Portion(fh, run.Text.Text));
            }
        }

        private int ParseLevel()
        {
            // gets default paragraph level font height for current paragraph's level
            var lvl = _aParagraph.ParagraphProperties?.Level?.Value;
            if (lvl == null)
            {
                lvl = 1;
            }
            else
            {
                // by unknown reason, slide and presentation's default settings have different numbering
                lvl++;
            }

            return (int)lvl;
        }

        #endregion Private Methods
    }
}
