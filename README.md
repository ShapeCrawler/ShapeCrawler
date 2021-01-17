<h3 align="center">

![ShapeCrawler](/resources/readme.png)

</h3>

<h3 align="center">

[![NuGet](https://img.shields.io/nuget/v/ShapeCrawler?color=blue)](https://www.nuget.org/packages/ShapeCrawler) [![.NET Standard](https://img.shields.io/badge/.NET%20Core-2.0-blue)](#) [![.NET Standard](https://img.shields.io/badge/.NET%20Standard-%3E%3D%202.0-blue.svg)](#) [![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE) 

</h3>

✅ **Project status: active**

ShapeCrawler (formerly SlideDotNet) is a .NET library for manipulating PowerPoint presentations. It provides fluent APIs to process slides without having Microsoft Office installed.

## Getting Started
You can quickly start work with the library by following steps listed below.
## Install

- [NuGet](https://nuget.org/packages/ShapeCrawler): `dotnet add package ShapeCrawler`

## Usage

### Text
```C#
using System;
using System.Collections.Generic;
using System.Linq;

using ShapeCrawler;
using ShapeCrawler.Texts;

public class TextSample
{
    public static void Text()
    {
        // Open presentation and get first slide
        PresentationSc presentation = PresentationSc.Open("helloWorld.pptx", isEditable: true);
        SlideSc slide = presentation.Slides.First();

        // Print text from all text shapes
        foreach (ShapeSc shape in slide.Shapes)
        {
            if (shape.HasTextFrame)
            {
                Console.WriteLine(shape.TextFrame.Text);
            }
        }

        // Change paragraph text
        ShapeSc textShape = slide.Shapes.First(sp => sp.HasTextFrame);
        ParagraphSc paragraph = textShape.TextFrame.Paragraphs.First();
        paragraph.Text = "A new paragraph text";

        // Print name and sizes of paragraph text portions
        ITextFrame textFrame = slide.Shapes.First(sp => sp.HasTextFrame).TextFrame;
        IEnumerable<Portion> paragraphPortions = textFrame.Paragraphs.First().Portions;
        foreach (Portion portion in paragraphPortions)
        {
            Console.WriteLine($"Font name: {portion.Font.Name}");
            Console.WriteLine($"Font size: {portion.Font.Size}");
        }

        // Save and close presentation
        presentation.Close();
    }
}
```

### Chart
```C#
using System;
using System.Linq;

using ShapeCrawler;
using ShapeCrawler.Charts;

public class ChartSample
{
    public static void Chart()
    {
        PresentationSc presentation = PresentationSc.Open("helloWorld.pptx", isEditable: false);
        SlideSc slide = presentation.Slides.First();

        // Get chart
        ShapeSc chartShape = slide.Shapes.First(sp => sp.HasChart == true);
        ChartSc chart = chartShape.Chart;
        
        // Print title string if the chart has a title
        if (chart.HasTitle)
        {
            Console.WriteLine(chart.Title);
        }
        
        if (chart.Type == ChartType.BarChart)
        {
            Console.WriteLine("Chart type is BarChart.");
        }
    }
}
```

### Slide Master
```C#
using ShapeCrawler;
using ShapeCrawler.SlideMaster;

public class SlideMasterSample
{
    public static void SlideMaster()
    {
        // Open presentation in the read mode
        PresentationSc presentation = PresentationSc.Open("helloWorld.pptx", isEditable: false);

        // Get number of Slide Masters in the presentation
        int slideMastersCount = presentation.SlideMasters.Count;

        // Get first Slide Master
        SlideMasterSc slideMaster = presentation.SlideMasters[0];

        // Get number of shapes in the Slide Master
        int masterShapeCount = slideMaster.Shapes.Count;
    }
}
```
# Feedback and Give a Star! :star:
The project is in development, and I’m pretty sure there are still lots of things to add in this library. Try it out and let me know your thoughts.

Feel free to submit a [ticket](https://github.com/ShapeCrawler/ShapeCrawler/issues) if you find bugs. Your valuable feedback is much appreciated to better improve this project. If you find this useful, please give it a star to show your support for this project. 

# Contributing
1. Fork it (https://github.com/ShapeCrawler/ShapeCrawler/fork)
2. Create your feature branch (`git checkout -b my-new-feature`) from master.
3. Commit your changes (`git commit -am 'Add some feature'`).
4. Push to the branch (`git push origin my-new-feature`).
5. Create a new Pull Request.

# Changelog
## Version 0.12.0 - 2021-01-17
### Added
- Added base API to get Slide Master collection — `PresentationSc.SlideMasters` (#44)
### Fixed
- Fixed New Line character processing for text paragraph (#87)

To find out more, please check out the [CHANGELOG](https://github.com/ShapeCrawler/ShapeCrawler/blob/master/CHANGELOG.md).
