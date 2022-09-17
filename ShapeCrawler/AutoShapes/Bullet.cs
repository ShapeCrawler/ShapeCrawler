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
        private readonly Lazy<SCBulletType> type;

        internal Bullet(A.ParagraphProperties aParagraphProperties)
        {
            this.aParagraphProperties = aParagraphProperties;
            this.type = new Lazy<SCBulletType>(this.ParseType);
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
        public SCBulletType Type => this.type.Value;

        #endregion Public Properties

        private SCBulletType ParseType()
        {
            if (this.aParagraphProperties == null)
            {
                return SCBulletType.None;
            }

            A.AutoNumberedBullet? aAutoNumeredBullet = this.aParagraphProperties.GetFirstChild<A.AutoNumberedBullet>();
            if (aAutoNumeredBullet != null)
            {
                return SCBulletType.Numbered;
            }

            A.PictureBullet? aPictureBullet = this.aParagraphProperties.GetFirstChild<A.PictureBullet>();
            if (aPictureBullet != null)
            {
                return SCBulletType.Picture;
            }

            A.CharacterBullet? aCharBullet = this.aParagraphProperties.GetFirstChild<A.CharacterBullet>();
            if (aCharBullet != null)
            {
                return SCBulletType.Character;
            }

            return SCBulletType.None;
        }

        private string? ParseColorHex()
        {
            if (this.Type == SCBulletType.None)
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

        private string? ParseChar()
        {
            if (this.Type == SCBulletType.None)
            {
                return null;
            }

            A.CharacterBullet? aCharBullet = this.aParagraphProperties.GetFirstChild<A.CharacterBullet>();
            if (aCharBullet == null)
            {
                throw new RuntimeDefinedPropertyException($"This is not {nameof(SCBulletType.Character)} type bullet.");
            }

            return aCharBullet.Char?.Value;
        }

        private string? ParseFontName()
        {
            if (this.Type == SCBulletType.None)
            {
                return null;
            }

            A.BulletFont? aBulletFont = this.aParagraphProperties.GetFirstChild<A.BulletFont>();
            return aBulletFont?.Typeface?.Value;
        }

        private int ParseSize()
        {
            if (this.Type == SCBulletType.None)
            {
                return 0;
            }

            A.BulletSizePercentage? aBulletSizePercent = this.aParagraphProperties.GetFirstChild<A.BulletSizePercentage>();
            int basicPoints = aBulletSizePercent?.Val?.Value ?? 100000;

            return basicPoints / 1000;
        }
    }
}