using System;

// ReSharper disable once CheckNamespace
#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Color.
/// </summary>
public struct Color
{
    /// <summary>
    ///     Predefined black color.
    /// </summary>
    public static readonly Color Black = new(0, 0, 0);

    /// <summary>
    ///     Predefined transparent color.
    /// </summary>
    public static readonly Color NoColor = new(0, 0, 0, 0);

    /// <summary>
    ///     Predefined white color.
    /// </summary>
    public static readonly Color White = new(255, 255, 255);

    /// <summary>
    ///     Max opacity value, equivalent to 1.
    /// </summary>
    internal const float Opacity = 255;

    private readonly int blue;
    private readonly int green;
    private readonly int red;

    private Color(int red, int green, int blue)
        : this(red, green, blue, 255)
    {
    }

    private Color(int red, int green, int blue, float alpha)
    {
        this.red = red;
        this.green = green;
        this.blue = blue;
        this.Alpha = alpha;
    }

    /// <summary>
    ///     Gets or sets the alpha value.
    /// </summary>
    /// <remarks>
    /// Values are 0 to 255, where 0 is totally transparent.
    /// </remarks>
    public float Alpha { get; set; }

    /// <summary>
    ///     Gets hexadecimal code.
    /// </summary>
    public string Hex => this.ToString();

    /// <summary>
    ///     Gets a value indicating whether the color is solid.
    /// </summary>
    internal readonly bool IsSolid => Math.Abs(this.Alpha - 255) < 0.01;

    /// <summary>
    ///     Gets a value indicating whether the color is transparent.
    /// </summary>
    internal readonly bool IsTransparent => Math.Abs(this.Alpha) < 0.01;

    /// <summary>
    ///     Creates color from Hex value.
    /// </summary>
    /// <param name="hex">Hex value.</param>
    /// <returns>Returns <see langword="true" /> if hex is a valid value. </returns>
    public static Color FromHex(string hex)
    {
        var value = hex.StartsWith("#", StringComparison.Ordinal) ? hex[1..] : hex;
        (int r, int g, int b, float a) = ParseHexValue(value);

        return new(r, g, b, a);
    }

    /// <summary>
    ///     Creates color hexadecimal code.
    /// </summary>
    public override string ToString() => $"{this.red:X2}{this.green:X2}{this.blue:X2}";

    /// <summary>
    ///     Returns a color of RGBA.
    /// </summary>
    /// <param name="hex">Hex value.</param>
    /// <returns>An RGBA color.</returns>
    private static (int, int, int, float) ParseHexValue(string hex)
    {
        if (string.IsNullOrEmpty(hex))
        {
            throw new ArgumentException("Hex value cannot be null or empty", nameof(hex));
        }

        return hex.Length switch
        {
            3 => ParseThreeDigitHex(hex),               // F00
            4 => ParseFourDigitHex(hex),                // FFFF
            6 => ParseSixDigitHex(hex),                 // FF0000
            8 => ParseEightDigitHex(hex),               // FFFFFF00
            _ => throw new FormatException("Hex value is invalid")
        };

        // Helper method to convert a hex character to integer value
        static int HexValue(char hex)
        {
            return Convert.ToInt32($"0x{hex}", 16);
        }

        // Parses 3-digit hex color (F00) -> (r,g,b,a)
        static (int, int, int, float) ParseThreeDigitHex(string hex)
        {
            int r = 17 * HexValue(hex[0]);
            int g = 17 * HexValue(hex[1]);
            int b = 17 * HexValue(hex[2]);
            return (r, g, b, 255); // Full opacity
        }

        // Parses 4-digit hex color (FFFF) -> (r,g,b,a)
        static (int, int, int, float) ParseFourDigitHex(string hex)
        {
            var rgbTuple = ParseThreeDigitHex(hex);
            int a = 17 * HexValue(hex[3]);
            return (rgbTuple.Item1, rgbTuple.Item2, rgbTuple.Item3, a);
        }

        // Parses 6-digit hex color (FF0000) -> (r,g,b,a)
        static (int, int, int, float) ParseSixDigitHex(string hex)
        {
            int r = (16 * HexValue(hex[0])) + HexValue(hex[1]);
            int g = (16 * HexValue(hex[2])) + HexValue(hex[3]);
            int b = (16 * HexValue(hex[4])) + HexValue(hex[5]);
            return (r, g, b, 255); // Full opacity
        }

        // Parses 8-digit hex color (FFFFFF00) -> (r,g,b,a) 
        static (int, int, int, float) ParseEightDigitHex(string hex)
        {
            var rgbTuple = ParseSixDigitHex(hex);
            int a = (16 * HexValue(hex[6])) + HexValue(hex[7]);
            return (rgbTuple.Item1, rgbTuple.Item2, rgbTuple.Item3, a);
        }
    }
}