namespace ShapeCrawler.Units;

internal readonly ref struct Inches
{
    private readonly decimal inches;

    internal Inches(decimal inches)
    {
        this.inches = inches;
    }

    internal decimal AsPixels() => this.inches * 96;
    
    internal float AsFloatPixels() => (float)this.inches * 96;
}