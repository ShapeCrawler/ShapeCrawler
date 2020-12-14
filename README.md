<h3 align="center">

![ShapeCrawler](/resources/readme.png)

</h3>

<h3 align="center">

  [![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
  [![NuGet](https://img.shields.io/nuget/v/SlideDotNet?color=blue)](https://www.nuget.org/packages/SlideDotNet)  

</h3>

ShapeCrawler (formerly SlideDotNet) is a fluent wrapper around [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK) for the processing of PowerPoint files without Microsoft Office installed. It aims to provide an intuitive and user-friendly interface to dealing with the underlying Open XML SDK API.

## Getting Started
You can quickly start work with the library by following steps listed below.

### Installing
To install ShapeCrawler, run the following command in the Package Manager Console:
```
PM> Install-Package ShapeCrawler
```

### Usage
```C#
public static async void Usage()
{
    // #1 Gets number of slides
    var presentation = new PresentationEx(@"c:\test.pptx");
    var slides = presentation.Slides;
    var numSlides = slides.Count();
    
    // #2 Gets number of shapes
    var firstSlide = slides[0];
    var shapes = firstSlide.Shapes;
    var numShapes = shapes.Count;
    
    // #3 Prints texts of shapes on the Debug console
    foreach (var sp in shapes)
    {
        if (sp.HasTextFrame)
        {
            Debug.WriteLine(sp.TextFrame.Text);
        }
    }

    // #4 Gets slide background content
    byte[] backgroundBytes = await firstSlide.BackgroundImage.GetImageBytes();
}
```

## Changelog
### Version 0.7.0 - 2020-10-12
#### Added
- Added `Bullet` property for the paragraph:
    - Bullet.Type
    - Bullet.Char
    - Bullet.FontName
    - Bullet.Size
    - Bullet.ColorHex

To find out more, please check out the [CHANGELOG](https://github.com/ShapeCrawler/ShapeCrawler/blob/master/CHANGELOG.md).

## Contribution
- Feel free to report a bug or suggest a new feature by creating an [issue](https://github.com/ShapeCrawler/ShapeCrawler/issues);
- Welcome to contribute. We are wating for your [Pull Requests](https://github.com/ShapeCrawler/ShapeCrawler/pulls). 
