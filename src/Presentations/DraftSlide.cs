﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ShapeCrawler.Presentations;

/// <summary>
///     Represents a draft for building a slide.
/// </summary>
public sealed class DraftSlide
{
    private readonly List<Action<ISlide>> actions = [];

    /// <summary>
    ///     Adds a picture to the slide with the specified name and geometry in points.
    /// </summary>
    public DraftSlide Picture(string name, int x, int y, int width, int height, Stream image)
    {
        this.actions.Add(slide =>
        {
            slide.Shapes.AddPicture(image);

            // Modify the last added picture
            var picture = slide.Shapes.Last();
            picture.Name = name;
            picture.X = x;
            picture.Y = y;
            picture.Width = width;
            picture.Height = height;
        });

        return this;
    }

    /// <summary>
    ///     Configures a picture using a nested builder.
    /// </summary>
    public DraftSlide Picture(Action<DraftPicture> configure)
    {
        this.actions.Add(slide =>
        {
            var b = new DraftPicture();
            configure(b);
            slide.Shapes.AddPicture(b.ImageStream);
            var pic = slide.Shapes.Last();
            pic.Name = b.DraftName;
            pic.X = b.DraftX;
            pic.Y = b.DraftY;
            pic.Width = b.DraftWidth;
            pic.Height = b.DraftHeight;
            if (!string.IsNullOrEmpty(b.GeometryName))
            {
                pic.GeometryType = (Geometry)Enum.Parse(typeof(Geometry), b.GeometryName!.Replace(" ", string.Empty));
            }
        });

        return this;
    }

    /// <summary>
    ///     Adds a text box (auto shape) and sets its content.
    /// </summary>
    public DraftSlide TextBox(string name, int x, int y, int width, int height, string content)
    {
        this.actions.Add(slide =>
        {
            slide.Shapes.AddShape(x, y, width, height, Geometry.Rectangle, content);
            var addedShape = slide.Shapes.Last<IShape>();
            addedShape.Name = name;
        });

        return this;
    }

    /// <summary>
    ///     Configures a text box using a nested builder.
    /// </summary>
    public DraftSlide TextBox(Action<DraftTextBox> configure)
    {
        this.actions.Add(slide =>
        {
            var builder = new DraftTextBox();
            configure(builder);
            slide.Shapes.AddShape(builder.PosX, builder.PosY, builder.BoxWidth, builder.BoxHeight, Geometry.Rectangle);
            var addedShape = slide.Shapes.Last<IShape>();
            addedShape.Name = builder.TextBoxName;
            if (!string.IsNullOrEmpty(builder.Content))
            {
                addedShape.TextBox!.SetText(builder.Content!);
            }
        });

        return this;
    }

    /// <summary>
    ///     Adds a line shape.
    /// </summary>
    public DraftSlide Line(string name, int startPointX, int startPointY, int endPointX, int endPointY)
    {
        this.actions.Add(slide =>
        {
            slide.Shapes.AddLine(startPointX, startPointY, endPointX, endPointY);
            var line = slide.Shapes.Last();
            line.Name = name;
        });

        return this;
    }

    /// <summary>
    ///     Adds a video shape and sets its properties.
    /// </summary>
    /// <param name="name">Requested shape name (ignored to keep a stable "Video" name as used by tests/examples).</param>
    /// <param name="x">X coordinate in points.</param>
    /// <param name="y">Y coordinate in points.</param>
    /// <param name="elementWidth">Width in points.</param>
    /// <param name="elementHeight">Height in points.</param>
    /// <param name="content">Video stream.</param>
    public DraftSlide Video(string name, int x, int y, int elementWidth, int elementHeight, Stream content)
    {
        this.actions.Add(slide =>
        {
            slide.Shapes.AddVideo(x, y, content);
            var media = slide.Shapes.Last();
            media.Name = name;
            media.X = x;
            media.Y = y;
            media.Width = elementWidth;
            media.Height = elementHeight;
        });

        return this;
    }

    /// <summary>
    ///     Adds a table with specified size.
    /// </summary>
    public DraftSlide Table(string name, int x, int y, int columnsCount, int rowsCount)
    {
        this.actions.Add(slide =>
        {
            slide.Shapes.AddTable(x, y, columnsCount, rowsCount);
            var table = slide.Shapes.Last<IShape>();
            table.Name = name;
        });

        return this;
    }

    /// <summary>
    ///     Adds a pie chart with specified name.
    /// </summary>
    public DraftSlide PieChart(string name)
    {
        this.actions.Add(slide =>
        {
            var categoryValues = new Dictionary<string, double>
            {
                { "Category 1", 40 },
                { "Category 2", 30 },
                { "Category 3", 30 }
            };
            slide.Shapes.AddPieChart(100, 100, 400, 300, categoryValues, "Series 1", name);
        });

        return this;
    }

    internal void ApplyTo(Presentation presentation)
    {
        // Always add a new slide for each DraftSlide application
        var sdkPres = presentation.PresDocument.PresentationPart!.Presentation;
        sdkPres.SlideIdList ??= new DocumentFormat.OpenXml.Presentation.SlideIdList();

        var blankLayout = presentation.SlideMasters[0].SlideLayouts.First(l => l.Name == "Blank");
        presentation.Slides.Add(blankLayout.Number);

        // Target the newly added slide
        var slide = presentation.Slides[presentation.Slides.Count - 1];
        foreach (var action in this.actions)
        {
            action(slide);
        }
    }
}