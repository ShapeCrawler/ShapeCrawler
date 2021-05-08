using System;
using System.Collections.Generic;
using System.Linq;
using ShapeCrawler.Exceptions;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.AutoShapes
{
    /// <summary>
    ///     Represents a paragraph bullet.
    /// </summary>
    public class Bullet
    {
        private readonly A.ParagraphProperties aParagraphProperties;
        private readonly Lazy<string> character;
        private readonly Lazy<string> colorHex;
        private readonly Lazy<string> fontName;
        private readonly Lazy<int> size;
        private readonly Lazy<BulletType> type;

        internal Bullet(A.ParagraphProperties aParagraphProperties)
        {
            this.aParagraphProperties = aParagraphProperties;
            this.type = new Lazy<BulletType>(this.ParseType);
            this.colorHex = new Lazy<string>(this.ParseColorHex);
            this.character = new Lazy<string>(this.ParseChar);
            this.fontName = new Lazy<string>(this.ParseFontName);
            this.size = new Lazy<int>(this.ParseSize);
        }

        #region Public Properties

        /// <summary>
        ///     Gets RGB color in HEX format.
        /// </summary>
        public string ColorHex => this.colorHex.Value;

        /// <summary>
        ///     Gets bullet character.
        /// </summary>
        public string Character => this.character.Value;

        /// <summary>
        ///     Gets bullet font name.
        /// </summary>
        public string FontName => this.fontName.Value;

        /// <summary>
        ///     Gets bullet size.
        /// </summary>
        public int Size => this.size.Value;

        /// <summary>
        ///     Gets bullet type.
        /// </summary>
        public BulletType Type => this.type.Value;

        #endregion Public Properties

        private BulletType ParseType()
        {
            if (this.aParagraphProperties == null)
            {
                return BulletType.None;
            }

            A.AutoNumberedBullet aAutoNumeredBullet = this.aParagraphProperties.GetFirstChild<A.AutoNumberedBullet>();
            if (aAutoNumeredBullet != null)
            {
                return BulletType.Numbered;
            }

            A.PictureBullet aPictureBullet = this.aParagraphProperties.GetFirstChild<A.PictureBullet>();
            if (aPictureBullet != null)
            {
                return BulletType.Picture;
            }

            A.CharacterBullet aCharBullet = this.aParagraphProperties.GetFirstChild<A.CharacterBullet>();
            if (aCharBullet != null)
            {
                return BulletType.Character;
            }

            return BulletType.None;
        }

        private string ParseColorHex()
        {
            if (this.Type == BulletType.None)
            {
                return null;
            }

            IEnumerable<A.RgbColorModelHex> aRgbClrModelHexCollection = this.aParagraphProperties.Descendants<A.RgbColorModelHex>();
            if (aRgbClrModelHexCollection.Any())
            {
                return aRgbClrModelHexCollection.Single().Val;
            }

            return null;
        }

        private string ParseChar()
        {
            if (this.Type == BulletType.None)
            {
                return null;
            }

            A.CharacterBullet aCharBullet = this.aParagraphProperties.GetFirstChild<A.CharacterBullet>();
            if (aCharBullet == null)
            {
                throw new RuntimeDefinedPropertyException($"This is not {nameof(BulletType.Character)} type bullet.");
            }

            return aCharBullet.Char.Value;
        }

        private string ParseFontName()
        {
            if (this.Type == BulletType.None)
            {
                return null;
            }

            A.BulletFont aBulletFont = this.aParagraphProperties.GetFirstChild<A.BulletFont>();
            return aBulletFont?.Typeface.Value;
        }

        private int ParseSize()
        {
            if (this.Type == BulletType.None)
            {
                return 0;
            }

            A.BulletSizePercentage aBulletSizePercent = this.aParagraphProperties.GetFirstChild<A.BulletSizePercentage>();
            int basicPoints = aBulletSizePercent?.Val.Value ?? 100000;

            return basicPoints / 1000;
        }
    }
}