using System;
using System.Collections.Generic;
using System.Linq;
using ShapeCrawler.Collections;
using ShapeCrawler.Models.Settings;
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

        private readonly A.Paragraph _aParagraph;
        private readonly Lazy<string> _text;
        private readonly Lazy<Bullet> _bullet;
        private readonly Lazy<PortionCollection> _portions;

        #endregion Fields

        #region Properties

        /// <summary>
        /// Gets or sets the the plain text of a paragraph.
        /// </summary>
        public string Text
        {
            get => _text.Value;
            set => SetText(value);
        }

        private void SetText(string text)
        {
            var firstPortion = Portions.First();
            
            throw new NotImplementedException();
        }

        /// <summary>
        /// Gets collection of paragraph text portions.
        /// </summary>
        public PortionCollection Portions => _portions.Value;

        /// <summary>
        /// Gets paragraph bullet. Returns null if bullet does not exist.
        /// </summary>
        public Bullet Bullet => _bullet.Value;

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initializes an instance of the <see cref="Paragraph"/> class.
        /// </summary>
        public Paragraph(IShapeContext spContext, A.Paragraph aParagraph)
        {
            _aParagraph = aParagraph;
            var innerPrLvl = GetInnerLevel(aParagraph);
            _text = new Lazy<string>(GetText);
            _bullet = new Lazy<Bullet>(GetBullet);
#if NETSTANDARD2_0
            _portions = new Lazy<PortionCollection>(()=>PortionCollection.Create(aParagraph, spContext, innerPrLvl, this));
#else
            _portions = new Lazy<PortionCollection>(PortionCollection.Create(aParagraph, spContext, innerPrLvl, this));
#endif
        }

        private Bullet GetBullet()
        {
            return new Bullet(_aParagraph.ParagraphProperties);
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

#endregion Private Methods
    }
}
