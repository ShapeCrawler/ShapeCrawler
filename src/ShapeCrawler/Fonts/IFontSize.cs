﻿namespace ShapeCrawler.Fonts;

/// <summary>
///     Represents a font size.
/// </summary>
internal interface IFontSize
{
    /// <summary>
    ///     Font size in points.
    /// </summary>
    decimal Size();
    
    /// <summary>
    ///     Updates font size in points.
    /// </summary>
    void Update(decimal points);
}