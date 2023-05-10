using System;

namespace ShapeCrawler.Drawing;

/// <summary>
/// Color.
/// </summary>
public class SCColor
{
    /// <summary>
    /// Gets a black color.
    /// </summary>
    public static readonly SCColor Black = new(0, 0, 0);

    /// <summary>
    /// Gets a transparent color.
    /// </summary>
    public static readonly SCColor Transparent = new(0, 0, 0, 0);

    /// <summary>
    /// Gets a white color.
    /// </summary>
    public static readonly SCColor White = new(255, 255, 255);

    /// <summary>
    /// Max opacity value, equivalent to 1.
    /// </summary>
    internal const float OPACITY = 255;

    /// <summary>
    /// Set color blue.
    /// </summary>
    private readonly int blue;

    /// <summary>
    /// Set color green.
    /// </summary>
    private readonly int green;

    /// <summary>
    /// Set color red.
    /// </summary>
    private readonly int red;

    /// <summary>
    /// Initializes a new instance of the <see cref="SCColor"/> class.
    /// </summary>
    /// <example>
    /// This shows how to creates a new instance of white color.
    /// <code>
    /// var white = new SCColor("FFF");
    /// var white = new SCColor("FFFF"); // Four char is for opacity (RGBA)
    /// var white = new SCColor("FFFFFF");
    /// var white = new SCColor("FFFFFF7F"); // 7F = ~50%
    /// </code>
    /// </example>
    /// <exception cref="Exception">Thrown when RGBA value is invalid.</exception>
    /// <remarks>
    /// "<paramref name="hex"/>" requires a RGBA value, but alpha (A) is optional.
    /// </remarks>
    /// <param name="hex">RGBA value.</param>
    public SCColor(string hex) 
        : this(ParseHexValue(hex))
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="SCColor"/> class.
    /// </summary>
    /// <param name="red">Red value.</param>
    /// <param name="green">Green value.</param>
    /// <param name="blue">Blue value.</param>
    public SCColor(int red, int green, int blue)
        : this(red, green, blue, 255)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="SCColor"/> class.
    /// </summary>
    /// <param name="red">Red value.</param>
    /// <param name="green">Green value.</param>
    /// <param name="blue">Blue value.</param>
    /// <param name="alpha">Alpha value.</param>
    public SCColor(int red, int green, int blue, float alpha)
    {
        this.red = red;
        this.green = green;
        this.blue = blue;
        this.Alpha = alpha;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="SCColor"/> class.
    /// </summary>
    /// <remarks>
    /// This constructor is reserved for <see cref="SCColor(string)"/>.
    /// </remarks>
    /// <param name="color">A RGBA tuple.</param>
    private SCColor((int r, int g, int b, float a) color) 
        : this(color.r, color.g, color.b, color.a)
    {
    }

    /// <summary>
    /// Gets or sets the alpha value.
    /// </summary>
    /// <remarks>
    /// Values are 0 to 255, where 0 is totally transparent.
    /// </remarks>
    public float Alpha { get; set; }

    /// <summary>
    /// Gets the blue value.
    /// </summary>
    public int B => this.blue;

    /// <summary>
    /// Gets the green value.
    /// </summary>
    public int G => this.green;

    /// <summary>
    /// Gets the red value.
    /// </summary>
    public int R => this.red;

    /// <summary>
    /// Gets a value indicating whether if color is solid.
    /// </summary>
    internal bool IsSolid => this.Alpha == 255;

    /// <summary>
    /// Gets a value indicating whether if color is transparent.
    /// </summary>
    internal bool IsTransparent => this.Alpha == 0;

    /// <summary>
    /// Returns a value indicating wheather hex is a valid value.
    /// </summary>
    /// <param name="hex">Hex value.</param>
    /// <param name="result">Color value.</param>
    /// <returns>Returns <see langword="true" /> if hex is a valid value. </returns>
    public static bool TryGetColorFromHex(string hex, out SCColor result)
    {
        result = SCColor.Black;

        try
        {
            // We can try to parse:
            // 3 or 6 chars without alpha (rgb): F01, FF0011,
            // 4 or 8 chars with alpha (rgba): F01F, FF0011FF 
            // Ignores hex values starting with "#" character.
            var value = hex.StartsWith("#", StringComparison.Ordinal) ? hex.Substring(1) : hex;

            // Parse value.
            (int r, int g, int b, float a) = ParseHexValue(value);

            // Creates a new instance
            result = new(r, g, b, a);

            return true;
        }
        catch (Exception)
        {
            return false;
        }
    }

    /// <inheritdoc/>
    public override string ToString()
    {
        // String representation ignores alpha value.
        return $"{this.R:X2}{this.G:X2}{this.B:X2}";
    }

    /// <summary>
    /// Returns a color of RGBA.
    /// </summary>
    /// <param name="hex">Hex value.</param>
    /// <returns>A RGBA color.</returns>
    private static (int, int, int, float) ParseHexValue(string hex)
    {
        int r;
        int b;
        int g;
        int a = 255;

        switch (hex.Length)
        {
            // FFFF 
            case 4:
                a = 17 * HexValue(hex[3]);
                goto case 3;
            case 3:
                // F00
                r = 17 * HexValue(hex[0]);
                g = 17 * HexValue(hex[1]);
                b = 17 * HexValue(hex[2]);
                break;
            case 8:
                // FFFFFF00
                a = (16 * HexValue(hex[6])) + HexValue(hex[7]);
                goto case 6;
            case 6:
                r = (16 * HexValue(hex[0])) + HexValue(hex[1]);
                g = (16 * HexValue(hex[2])) + HexValue(hex[3]);
                b = (16 * HexValue(hex[4])) + HexValue(hex[5]);
                break;
            default:
                // String format is invalid.
                throw new FormatException("Hex value is invalid");
        }

        static int HexValue(char hex)
        {
            return Convert.ToInt32($"0x{hex}", 16);
        }

        return (r, g, b, a);
    }
}
