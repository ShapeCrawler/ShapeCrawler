using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Drawing;
using ShapeCrawler.Exceptions;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a paragraph bullet.
/// </summary>
public sealed class Bullet // TODO: extract interface
{
    private readonly A.ParagraphProperties aParagraphProperties;
    private readonly Lazy<string?> character;
    private readonly Lazy<string?> colorHex;
    private readonly Lazy<string?> fontName;
    private readonly Lazy<int> size;
    private readonly Lazy<BulletType> type;

    internal Bullet(A.ParagraphProperties aParagraphProperties)
    {
        this.aParagraphProperties = aParagraphProperties;
        this.type = new Lazy<BulletType>(this.ParseType);
        this.colorHex = new Lazy<string?>(this.ParseColorHex);
        this.character = new Lazy<string?>(this.ParseChar);
        this.fontName = new Lazy<string?>(this.ParseFontName);
        this.size = new Lazy<int>(this.ParseSize);
    }

    #region Public Properties

    /// <summary>
    ///     Gets RGB color in HEX format.
    /// </summary>
    public string? ColorHex => this.colorHex.Value;

    /// <summary>
    ///     Gets or sets bullet character.
    /// </summary>
    public string? Character
    {
        get => this.character.Value;

        set
        {
            if (this.aParagraphProperties == null || this.Type != BulletType.Character)
            {
                return;
            }

            A.CharacterBullet? aCharBullet = this.aParagraphProperties.GetFirstChild<A.CharacterBullet>();
            if (aCharBullet == null)
            {
                aCharBullet = new CharacterBullet();
                this.aParagraphProperties.AddChild(aCharBullet);
            }

            aCharBullet.Char = value;
        }
    }

    /// <summary>
    ///     Gets or sets bullet font name. Returns <see langword="null"/> if bullet doesn't exist.
    /// </summary>
    public string? FontName
    {
        get => this.fontName.Value;

        set
        {
            if (this.aParagraphProperties == null || this.Type == BulletType.None)
            {
                return;
            }

            A.BulletFont? aBulletFont = this.aParagraphProperties.GetFirstChild<A.BulletFont>();
            if (aBulletFont == null)
            {
                aBulletFont = new BulletFont();
                this.aParagraphProperties.AddChild(aBulletFont);
            }

            aBulletFont.Typeface = value;
        }
    }

    /// <summary>
    ///     Gets or sets bullet size in percentages of text.
    /// </summary>
    public int Size
    {
        get => this.size.Value;
        set
        {
            if (this.aParagraphProperties == null)
            {
                return;
            }

            A.BulletSizePercentage? aBulletSizePercent = this.aParagraphProperties.GetFirstChild<A.BulletSizePercentage>();
            if (aBulletSizePercent == null)
            {
                aBulletSizePercent = new A.BulletSizePercentage();
                this.aParagraphProperties.AddChild(aBulletSizePercent);
            }

            aBulletSizePercent.Val = value * 1000;
        }
    }

    /// <summary>
    ///     Gets or sets bullet type.
    /// </summary>
    public BulletType Type
    {
        get => this.type.Value;
        set
        {
            if (this.aParagraphProperties == null)
            {
                return;
            }

            A.AutoNumberedBullet? aAutoNumeredBullet = this.aParagraphProperties.GetFirstChild<A.AutoNumberedBullet>();
            this.aParagraphProperties.RemoveChild(aAutoNumeredBullet);

            A.PictureBullet? aPictureBullet = this.aParagraphProperties.GetFirstChild<A.PictureBullet>();
            this.aParagraphProperties.RemoveChild(aPictureBullet);

            A.CharacterBullet? aCharBullet = this.aParagraphProperties.GetFirstChild<A.CharacterBullet>();
            this.aParagraphProperties.RemoveChild(aCharBullet);

            if (value == BulletType.Numbered)
            {
                var child = new AutoNumberedBullet
                {
                    // replace at property
                    Type = A.TextAutoNumberSchemeValues.ArabicPeriod
                };

                this.aParagraphProperties.AddChild(child);
            }

            if (value == BulletType.Picture)
            {
                var child = new PictureBullet();
                this.aParagraphProperties.AddChild(child);
            }

            if (value == BulletType.Character)
            {
                var child = new CharacterBullet();
                this.aParagraphProperties.AddChild(child);
            }
        }
    }

    #endregion Public Properties

    private BulletType ParseType()
    {
        if (this.aParagraphProperties == null)
        {
            return BulletType.None;
        }

        A.AutoNumberedBullet? aAutoNumeredBullet = this.aParagraphProperties.GetFirstChild<A.AutoNumberedBullet>();
        if (aAutoNumeredBullet != null)
        {
            return BulletType.Numbered;
        }

        A.PictureBullet? aPictureBullet = this.aParagraphProperties.GetFirstChild<A.PictureBullet>();
        if (aPictureBullet != null)
        {
            return BulletType.Picture;
        }

        A.CharacterBullet? aCharBullet = this.aParagraphProperties.GetFirstChild<A.CharacterBullet>();
        if (aCharBullet != null)
        {
            return BulletType.Character;
        }

        return BulletType.None;
    }

    private string? ParseColorHex()
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

    private string? ParseChar()
    {
        if (this.Type == BulletType.None)
        {
            return null;
        }

        A.CharacterBullet? aCharBullet = this.aParagraphProperties.GetFirstChild<A.CharacterBullet>();
        if (aCharBullet == null)
        {
            throw new RuntimeDefinedPropertyException($"This is not {nameof(BulletType.Character)} type bullet.");
        }

        return aCharBullet.Char?.Value;
    }

    private string? ParseFontName()
    {
        if (this.Type == BulletType.None)
        {
            return null;
        }

        A.BulletFont? aBulletFont = this.aParagraphProperties.GetFirstChild<A.BulletFont>();
        return aBulletFont?.Typeface?.Value;
    }

    private int ParseSize()
    {
        if (this.Type == BulletType.None)
        {
            return 0;
        }

        A.BulletSizePercentage? aBulletSizePercent = this.aParagraphProperties.GetFirstChild<A.BulletSizePercentage>();
        int basicPoints = aBulletSizePercent?.Val?.Value ?? 100000;

        return basicPoints / 1000;
    }
}