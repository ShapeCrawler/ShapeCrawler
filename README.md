<h3 align="center">

![ShapeCrawler](./resources/readme.png)

</h3>

<h3 align="center">

[![NuGet](https://img.shields.io/nuget/v/ShapeCrawler?color=orange)](https://www.nuget.org/packages/ShapeCrawler) ![Nuget](https://img.shields.io/nuget/dt/ShapeCrawler?color=orange) [![License](https://img.shields.io/badge/license-MIT-orange.svg)](LICENSE) 

</h3>

✅ **Project status: active**

ShapeCrawler (formerly SlideDotNet) is a .NET library for manipulating PowerPoint presentations. It provides fluent APIs to process slides without having Microsoft Office installed.

This library provides a simplified object model on top of the [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK) for manipulating PowerPoint documents.

## Getting Started
Support targets: .NET Framework 4.6.1+, .NET Core 2.0+ and .NET 5+

To get started, install ShapeCrawler from [NuGet](https://nuget.org/packages/ShapeCrawler):
```console
dotnet add package ShapeCrawler
```

## Usage
The usage samples below will take you through some work experience with the presentation object model.

### Working with Texts
```C#
using System;
using System.Linq;
using ShapeCrawler;

public class TextSample
{
    public static void Text()
    {
        // Open presentation and get first slide
        using IPresentation presentation = SCPresentation.Open("helloWorld.pptx", isEditable: true);
        ISlide slide = presentation.Slides.First();

        // Get text holder auto shape
        IAutoShape autoShape = (IAutoShape)slide.Shapes.First(sp => sp is IAutoShape);

        // Change whole shape text
        autoShape.TextBox.Text = "A new shape text";

        // Change text for a certain paragraph
        IParagraph paragraph = autoShape.TextBox.Paragraphs[1];
        paragraph.Text = "A new text for second paragraph";

        // Get font name and size
        IPortion paragraphPortion = autoShape.TextBox.Paragraphs.First().Portions.First();
        Console.WriteLine($"Font name: {paragraphPortion.Font.Name}");
        Console.WriteLine($"Font size: {paragraphPortion.Font.Size}");

        // Set bold font
        paragraphPortion.Font.IsBold = true;

        // Get font ARGB color
        Color fontColor = paragraphPortion.Font.ColorFormat.Color;

        // Save changes presentation
        presentation.Save();
    }
}
```

### Working with Tables
```C#
using System;
using System.Linq;
using ShapeCrawler;

public class TableSample
{
    public static void Table()
    {
        // Get first slide
        using IPresentation presentation = SCPresentation.Open("helloWorld.pptx", isEditable: false);
        ISlide slide = presentation.Slides.First();

        // Get table
        ITable table = (ITable)slide.Shapes.First(sp => sp is ITable);

        // Get number of rows in the table
        int rowsCount = table.Rows.Count;

        // Get number of cells in the first row
        int rowCellsCount = table.Rows[0].Cells.Count;

        // Print a message if the cell is a part of a merged cells group
        foreach (SCTableRow row in table.Rows)
        {
            foreach (ITableCell cellItem in row.Cells)
            {
                if (cellItem.IsMergedCell)
                {
                    Console.WriteLine("The cell is a part of a merged cells group.");
                }
            }
        }

        // Get column's width
        Column tableColumn = table.Columns[0];
        long columnWidth = tableColumn.Width;

        // Get row's height
        long rowHeight = table.Rows[0].Height;

        // Get cell with row index 0 and column index 1
        ITableCell cell = table[0, 1];

        // Merge cells
        table.MergeCells(table[0,0], table[0, 1]);

        // Save changes
        presentation.Save();
    }
}
```

### Working with Pictures
```C#
using System.Linq;
using ShapeCrawler;

public class PictureSamples
{
    public static void Picture()
    {
        using IPresentation presentation = SCPresentation.Open("helloWorld.pptx", isEditable: true);

        // Get picture shape
        IPicture picture = presentation.Slides[0].Shapes.OfType<IPicture>().First();

        // Change image
        picture.Image.SetImage("new-image.png");

        // Save changes
        presentation.Save();
    }
}
```

### Working with Audio
```C#
using System.IO;
using ShapeCrawler;

public class WorkingWithAudio
{
    public static void AddAudioShape()
    {
        // Add audio
        IPresentation presentation = SCPresentation.Open("helloWorld.pptx", isEditable: true);                
        IShapeCollection shapes = presentation.Slides[0].Shapes;
        using Stream mp3Stream = File.OpenRead("audio.mp3");
        IAudioShape addedAudioShape = shapes.AddNewAudio(xPixel: 300, yPixels: 100, mp3Stream);

        presentation.Save();
        presentation.Close();      
    }
}
```

### Working with Charts
```C#
using System;
using System.Linq;
using ShapeCrawler;

public class ChartSample
{
    public static void Chart()
    {
        using IPresentation presentation = SCPresentation.Open("helloWorld.pptx", isEditable: true);
        ISlide slide = presentation.Slides.First();

        // Get chart
        IChart chart = (IChart)slide.Shapes.First(sp => sp is IChart);
        
        // Print chart title
        if (chart.HasTitle)
        {
            Console.WriteLine(chart.Title);
        }
        
        if (chart.Type == ChartType.BarChart)
        {
            Console.WriteLine("Chart type is BarChart.");
        }

        // Update category
        chart.Categories[0].Name = "Price";       
    }
}
```

### Working with Slide Masters

```C#
using ShapeCrawler;

public class SlideMasterSample
{
    public static void SlideMaster()
    {
        // Open presentation in the read mode
        using IPresentation presentation = SCPresentation.Open("helloWorld.pptx", isEditable: false);

        // Get number of Slide Masters in the presentation
        int slideMastersCount = presentation.SlideMasters.Count;

        // Get first Slide Master
        ISlideMaster slideMaster = presentation.SlideMasters[0];

        // Get number of shapes in the Slide Master
        int masterShapeCount = slideMaster.Shapes.Count;
    }
}
```

### Update slide collection

```C#
using System.Linq;
using ShapeCrawler;

public class UpdateSlideCollection
{
    public static void Update()
    {
        // Remove first slide
        using IPresentation presentation = SCPresentation.Open("test.pptx", true);
        ISlide removingSlide = presentation.Slides.First();
        presentation.Slides.Remove(removingSlide);

        // Move second slide to first position
        presentation.Slides[1].Number = 2;

        // Copy second slide from source into dest
        using IPresentation source = SCPresentation.Open("source.pptx", true);
        using IPresentation dest = SCPresentation.Open("dest.pptx", true);
        ISlide copyingSlide = presentation.Slides[1];
        dest.Slides.Add(copyingSlide);

        // Save changes
        presentation.Save();
    }
}
```

# Feedback and Give a Star! :star:
The project is in development, and I’m pretty sure there are still lots of things to add in this library. Try it out and let me know your thoughts.

Feel free to submit a [ticket](https://github.com/ShapeCrawler/ShapeCrawler/issues) if you find bugs. Your valuable feedback is much appreciated to improve this project better. If you find this useful, please give it a star to show your support. 

# Contributing and Support
1. Fork it (https://github.com/ShapeCrawler/ShapeCrawler/fork)
2. Create your feature branch (`git checkout -b my-new-feature`) from *master*.
3. Commit your changes (`git commit -am 'Add some feature'`).
4. Push to the branch (`git push origin my-new-feature`).
5. Create a new Pull Request.

Don't hesitate to contact me (theadamo86@gmail.com) if you have a questions or wanna be involved!
