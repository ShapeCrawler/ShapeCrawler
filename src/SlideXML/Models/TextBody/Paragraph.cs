using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using SlideXML.Models.Settings;
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
        private readonly ElementSettings _elSetting;
        private string _text;
        private readonly Lazy<int> _lvl;
        private Lazy<List<Portion>> _portions;

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
        public IList<Portion> Portions => _portions.Value;

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initializes an instance of the <see cref="Paragraph"/> class.
        /// </summary>
        /// <param name="elSetting"></param>
        /// <param name="aParagraph">A XML paragraph which contains a text.</param>
        public Paragraph(ElementSettings elSetting, A.Paragraph aParagraph)
        {
            _aParagraph = aParagraph ?? throw new ArgumentNullException(nameof(aParagraph));
            _elSetting = elSetting ?? throw new ArgumentNullException(nameof(elSetting));
            _lvl = new Lazy<int>(GetLevel(_aParagraph));
            _portions = new Lazy<List<Portion>>(GetPortions);
        }

        #endregion Constructors

        private static int GetLevel(A.Paragraph aPr)
        {
            var lvl = aPr.ParagraphProperties?.Level;
            if (lvl == null) 
            {
                return 1; // null is first level
            }

            return ++lvl.Value;
        }

        #region Private Methods

        private void InitText()
        {
            _text = Portions.Select(p => p.Text).Aggregate((t1, t2) => t1 + t2);
        }

        [SuppressMessage("ReSharper", "ConvertIfStatementToConditionalTernaryExpression")]
        private List<Portion> GetPortions()
        {
            var runs = _aParagraph.Elements<A.Run>();
            var portions = new List<Portion>(runs.Count());
            foreach (var run in runs)
            {
                var runFh = GetRunFontHeight(run);
                portions.Add(new Portion(runFh, run.Text.Text));
            }

            return portions;
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
            var ph = _elSetting.Placeholder;
            if (ph != null)
            {
                return ph.FontHeights[_lvl.Value];
            }

            // from global presentation setting
            return _elSetting.PreSettings.LlvFontHeights[_lvl.Value];
        }

        #endregion Private Methods
    }
}
