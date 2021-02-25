using System;
using System.Collections.Generic;
using System.Linq;
using ShapeCrawler.Exceptions;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.AutoShapes
{
    /// <summary>
    ///     Represents a paragraph bullet class.
    /// </summary>
    public class Bullet
    {
        private readonly A.ParagraphProperties _aParagraphProperties;
        private readonly Lazy<string> _char;
        private readonly Lazy<string> _colorHex;
        private readonly Lazy<string> _fontName;
        private readonly Lazy<int> _size;
        private readonly Lazy<BulletType> _type;

        #region Constructors

        internal Bullet(A.ParagraphProperties aParagraphProperties)
        {
            _aParagraphProperties = aParagraphProperties;
            _type = new Lazy<BulletType>(ParseType);
            _colorHex = new Lazy<string>(ParseColorHex);
            _char = new Lazy<string>(ParseChar);
            _fontName = new Lazy<string>(ParseFontName);
            _size = new Lazy<int>(ParseSize);
        }

        #endregion Constructors

        #region Properties

        public string ColorHex => _colorHex.Value;

        public string Char => _char.Value;

        public string FontName => _fontName.Value;

        /// <summary>
        ///     Gets bullet size, in percent of the text.
        /// </summary>
        public int Size => _size.Value;

        public BulletType Type => _type.Value;

        #endregion Properties

        #region Private Methods

        private BulletType ParseType()
        {
            if (_aParagraphProperties == null)
            {
                return BulletType.None;
            }

            var buAutoNum = _aParagraphProperties.GetFirstChild<A.AutoNumberedBullet>();
            if (buAutoNum != null)
            {
                return BulletType.Numbered;
            }

            var buBlip = _aParagraphProperties.GetFirstChild<A.PictureBullet>();
            if (buBlip != null)
            {
                return BulletType.Picture;
            }

            var buChar = _aParagraphProperties.GetFirstChild<A.CharacterBullet>();
            if (buChar != null)
            {
                return BulletType.Character;
            }

            return BulletType.None;
        }

        private string ParseColorHex()
        {
            if (Type == BulletType.None)
            {
                return null;
            }

            IEnumerable<A.RgbColorModelHex> aRgbClrModelHexCollection =
                _aParagraphProperties.Descendants<A.RgbColorModelHex>();
            if (aRgbClrModelHexCollection.Any())
            {
                return aRgbClrModelHexCollection.Single().Val;
            }

            return null;
        }

        private string ParseChar()
        {
            if (Type == BulletType.None)
            {
                return null;
            }

            var buChar = _aParagraphProperties.GetFirstChild<A.CharacterBullet>();
            if (buChar == null)
            {
                throw new RuntimeDefinedPropertyException($"This is not {nameof(BulletType.Character)} type bullet.");
            }

            return buChar.Char.Value;
        }

        private string ParseFontName()
        {
            if (Type == BulletType.None)
            {
                return null;
            }

            var buFont = _aParagraphProperties.GetFirstChild<A.BulletFont>();
            return buFont?.Typeface.Value;
        }

        private int ParseSize()
        {
            if (Type == BulletType.None)
            {
                return 0;
            }

            var buSzPct = _aParagraphProperties.GetFirstChild<A.BulletSizePercentage>();
            var basicPoints = buSzPct?.Val.Value ?? 100000;

            return basicPoints / 1000;
        }

        #endregion Private Methods
    }
}