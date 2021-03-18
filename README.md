<h3 align="center">

![ShapeCrawler](./resources/readme.png)

</h3>

<h3 align="center">

[![NuGet](https://img.shields.io/nuget/v/ShapeCrawler?color=orange)](https://www.nuget.org/packages/ShapeCrawler) ![Nuget](https://img.shields.io/nuget/dt/ShapeCrawler?color=orange) [![.NET Standard](https://img.shields.io/badge/.NET%20Core-5.0-orange)](#) [![.NET Standard](https://img.shields.io/badge/.NET%20Core-2.1-orange)](#) [![.NET Standard](https://img.shields.io/badge/.NET%20Standard-%3E%3D%202.0-orange.svg)](#) [![License](https://img.shields.io/badge/license-MIT-orange.svg)](LICENSE) 

</h3>

✅ **Project status: active**

ShapeCrawler (formerly SlideDotNet) is a .NET library for manipulating PowerPoint presentations. It provides fluent APIs to process slides without having Microsoft Office installed.

This library provides a simplified object model on top of the [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK) for manipulating PowerPoint documents.

## Getting Started
To get started use ShapeCrawler APIs install it from [NuGet](https://nuget.org/packages/ShapeCrawler):
```console
dotnet add package ShapeCrawler
```

## Usage
The usage samples below will take you through some work experience with the presentation object model.

### Working with Texts
```C#
using System;
using System.Collections.Generic;
using System.Linq;
using ShapeCrawler;
using ShapeCrawler.AutoShapes;

public class TextSample
{
    public static void Text()
    {
        // Open presentation and get first slide
        using PresentationSc presentation = PresentationSc.Open("helloWorld.pptx", isEditable: true);
        SlideSc slide = presentation.Slides.First();

        // Get text holder auto shape
        IAutoShape autoShape = (IAutoShape)slide.Shapes.First(sp => sp is IAutoShape);

        // Change whole shape text
        autoShape.TextBox.Text = "A new shape text";

        // Change text for a certain paragraph
        ParagraphSc paragraph = autoShape.TextBox.Paragraphs[1];
        paragraph.Text = "A new text for second paragraph";

        // Print font name and size of a paragraph text portions
        ITextBox textBox = autoShape.TextBox;
        IEnumerable<Portion> paragraphPortions = textBox.Paragraphs.First().Portions;
        foreach (Portion portion in paragraphPortions)
        {
            Console.WriteLine($"Font name: {portion.Font.Name}");
            Console.WriteLine($"Font size: {portion.Font.Size}");
        }

        // Save and close the presentation
        presentation.Close();
    }
}
```

### Working with Tables
```C#
using System;
using System.Linq;
using ShapeCrawler;
using ShapeCrawler.Tables;

public class TableSample
{
    public static void Table()
    {
        // Get first slide
        using PresentationSc presentation = PresentationSc.Open("helloWorld.pptx", isEditable: false);
        SlideSc slide = presentation.Slides.First();

        // Get table
        ITable table = (ITable)slide.Shapes.First(sp => sp is ITable);

        // Get number of rows in the table
        int rowsCount = table.Rows.Count;

        // Get number of cells in the first row
        int rowCellsCount = table.Rows[0].Cells.Count;

        // Print a message if the cell is a part of a merged cells group
        foreach (RowSc row in table.Rows)
        {
            foreach (CellSc cellItem in row.Cells)
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
        CellSc cell = table[0, 1];

        // Merge cells
        table.MergeCells(table[0,0], table[0, 1]);

        presentation.Close();
    }
}
```

### Working with Charts
```C#
using System;
using System.Linq;
using ShapeCrawler;
using ShapeCrawler.Charts;

public class ChartSample
{
    public static void Chart()
    {
        using PresentationSc presentation = PresentationSc.Open("helloWorld.pptx", isEditable: false);
        SlideSc slide = presentation.Slides.First();

        // Get chart
        IChart chart = (IChart)slide.Shapes.First(sp => sp is IChart);
        
        // Print title string if the chart has a title
        if (chart.HasTitle)
        {
            Console.WriteLine(chart.Title);
        }
        
        if (chart.Type == ChartType.BarChart)
        {
            Console.WriteLine("Chart type is BarChart.");
        }

        presentation.Close();
    }
}
```

### Working with Slide Masters

```C#
using ShapeCrawler;
using ShapeCrawler.SlideMaster;

public class SlideMasterSample
{
    public static void SlideMaster()
    {
        // Open presentation in the read mode
        using PresentationSc presentation = PresentationSc.Open("helloWorld.pptx", isEditable: false);

        // Get number of Slide Masters in the presentation
        int slideMastersCount = presentation.SlideMasters.Count;

        // Get first Slide Master
        SlideMasterSc slideMaster = presentation.SlideMasters[0];

        // Get number of shapes in the Slide Master
        int masterShapeCount = slideMaster.Shapes.Count;

        presentation.Close();
    }
}
```

# Known Issue
**Font Size** is a tricky part of PowerPoint document structure since obtaining this value leads to parsing different presentation layers —  Slide, Slide Layout or Slide Master. Hence, If you note that font size was incorrect defined, please report [an issue](https://github.com/ShapeCrawler/ShapeCrawler/issues) with attaching your pptx-file example.

# Feedback and Give a Star! :star:
The project is in development, and I’m pretty sure there are still lots of things to add in this library. Try it out and let me know your thoughts.

Feel free to submit a [ticket](https://github.com/ShapeCrawler/ShapeCrawler/issues) if you find bugs. Your valuable feedback is much appreciated to improve this project better. If you find this useful, please give it a star to show your support. 

# Contributing
1. Fork it (https://github.com/ShapeCrawler/ShapeCrawler/fork)
2. Create your feature branch (`git checkout -b my-new-feature`) from *master*.
3. Commit your changes (`git commit -am 'Add some feature'`).
4. Push to the branch (`git push origin my-new-feature`).
5. Create a new Pull Request.

Don't hesitate to contact me if you want to get involved!

# Changelog
## Version 0.16.1 - 2021-03-08
### Fixed
- Fixed parser of font properties

To find out more, please check out the [CHANGELOG](https://github.com/ShapeCrawler/ShapeCrawler/blob/master/CHANGELOG.md).
