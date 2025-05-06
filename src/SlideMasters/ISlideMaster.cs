using System;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <summary>
///     Represents a Slide Master.
/// </summary>
public interface ISlideMaster
{
    /// <summary>
    ///     Gets background image if slide master has background, otherwise <see langword="null"/>.
    /// </summary>
    IImage? Background { get; }

    /// <summary>
    ///     Gets slide layout collection.
    /// </summary>
    ISlideLayoutCollection SlideLayouts { get; }

    /// <summary>
    ///     Gets the collection of master shapes.
    /// </summary>
    IShapeCollection Shapes { get; }

    /// <summary>
    ///     Gets theme.
    /// </summary>
    ITheme Theme { get; }

    /// <summary>
    ///     Gets slide number. Returns <see langword="null"/> if slide master does not have slide number.
    /// </summary>
    IMasterSlideNumber? SlideNumber { get; }

    /// <summary>
    ///     Returns a shape from the slide master.
    /// </summary>
    /// <param name="shape">The name of the shape.</param>
    /// <returns>The requested shape.</returns>
    IShape Shape(string shape);

    /// <summary>
    ///     Gets slide layout by name.
    /// </summary>
    ISlideLayout SlideLayout(string name);

    /// <summary>
    ///     Gets slide layout by number.
    /// </summary>
    ISlideLayout SlideLayout(int number);
}

internal sealed class SlideMaster : ISlideMaster
{
    private readonly SlideLayoutCollection layouts;
    private readonly Lazy<MasterSlideNumber?> slideNumber;
    private readonly SlideMasterPart slideMasterPart;

    internal SlideMaster(SlideMasterPart slideMasterPart)
    {
        this.slideMasterPart = slideMasterPart;
        this.layouts = new SlideLayoutCollection(slideMasterPart);
        this.slideNumber = new Lazy<MasterSlideNumber?>(this.CreateSlideNumber);
        this.Shapes = new ShapeCollection(this.slideMasterPart);
    }

    public IImage? Background => null;

    public ISlideLayoutCollection SlideLayouts => this.layouts;

    public IShapeCollection Shapes { get; }

    public ITheme Theme => new Theme(this.slideMasterPart, this.slideMasterPart.ThemePart!.Theme);

    public IMasterSlideNumber? SlideNumber => this.slideNumber.Value;

    public int Number
    {
        get
        {
            var match = Regex.Match(this.slideMasterPart.Uri.ToString(), @"\d+", RegexOptions.None, TimeSpan.FromSeconds(1));
            return int.Parse(match.Value);      
        }
    } 

    public IShape Shape(string shape) => this.Shapes.Shape(shape);

    public ISlideLayout SlideLayout(string name) => this.layouts.First(l => l.Name == name);

    public ISlideLayout SlideLayout(int number) => this.layouts.First(l => l.Number == number);

    internal SlideLayout InternalSlideLayout(int number) => this.layouts.Layout(number);

    private MasterSlideNumber? CreateSlideNumber()
    {
        var pSldNum = this.slideMasterPart.SlideMaster.CommonSlideData!.ShapeTree!
            .Elements<P.Shape>()
            .FirstOrDefault(s =>
                s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value ==
                P.PlaceholderValues.SlideNumber);

        return pSldNum is null ? null : new MasterSlideNumber(pSldNum);
    }
}