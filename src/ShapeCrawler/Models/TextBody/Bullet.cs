using System;
using System.Linq;
using ShapeCrawler.Enums;
using ShapeCrawler.Exceptions;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Models.TextBody
{
    /// <summary>
    /// Represents a paragraph bullet class.
    /// </summary>
    public class Bullet
    {
        private static readonly string BulletIsNotDefinedErrorMsg = $"This bullet type is {BulletType.None}";
        private readonly A.ParagraphProperties _xmlParagraphProperties;
        private readonly Lazy<string> _colorHex;
        private readonly Lazy<string> _char;
        private readonly Lazy<string> _fontName;
        private readonly Lazy<int> _size;
        private readonly Lazy<BulletType> _type;

        #region Properties

        public string ColorHex => _colorHex.Value;

        public string Char => _char.Value;

        public string FontName => _fontName.Value;

        /// <summary>
        /// Gets bullet size, in percent of the text.
        /// </summary>
        public int Size => _size.Value;

        public BulletType Type => _type.Value;

        #endregion Properties

        #region Constructors

        public Bullet(A.ParagraphProperties xmlParagraphProperties)
        {
            _xmlParagraphProperties = xmlParagraphProperties;
            _type = new Lazy<BulletType>(ParseType);
            _colorHex = new Lazy<string>(ParseColorHex);
            _char = new Lazy<string>(ParseChar);
            _fontName = new Lazy<string>(ParseFontName);
            _size = new Lazy<int>(ParseSize);
        }

        #endregion Constructors

        #region Private Methods

        private BulletType ParseType()
        {
            if (_xmlParagraphProperties == null)
            {
                return BulletType.None;
            }
            var buAutoNum = _xmlParagraphProperties.GetFirstChild<A.AutoNumberedBullet>();
            if (buAutoNum != null)
            {
                return BulletType.Numbered;
            }
            var buBlip = _xmlParagraphProperties.GetFirstChild<A.PictureBullet>();
            if (buBlip != null)
            {
                return BulletType.Picture;
            }
            var buChar = _xmlParagraphProperties.GetFirstChild<A.CharacterBullet>();
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
                throw new RuntimeDefinedPropertyException(BulletIsNotDefinedErrorMsg);
            }

            var srgbClrElements = _xmlParagraphProperties.Descendants<A.RgbColorModelHex>();
            if (srgbClrElements.Any())
            {
                return srgbClrElements.Single().Val;
            }

            return null;
        }

        private string ParseChar()
        {
            if (Type == BulletType.None)
            {
                throw new RuntimeDefinedPropertyException(BulletIsNotDefinedErrorMsg);
            }

            var buChar = _xmlParagraphProperties.GetFirstChild<A.CharacterBullet>();
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
                throw new RuntimeDefinedPropertyException(BulletIsNotDefinedErrorMsg);
            }

            var buFont = _xmlParagraphProperties.GetFirstChild<A.BulletFont>();
            return buFont?.Typeface.Value;
        }

        private int ParseSize()
        {
            if (Type == BulletType.None)
            {
                throw new RuntimeDefinedPropertyException(BulletIsNotDefinedErrorMsg);
            }

            var buSzPct = _xmlParagraphProperties.GetFirstChild<A.BulletSizePercentage>();
            var basicPoints = buSzPct?.Val.Value ?? 100000;

            return basicPoints / 1000;
        }

        #endregion Private Methods
    }
}