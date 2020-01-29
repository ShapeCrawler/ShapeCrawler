![Alt text](/resources/readme.png)


SlideXML is a lightweight .NET library for parse Microsoft PowerPoint file presentations without having to install the PowerPoint application. It aims to provide an intuitive and user-friendly interface to dealing with the underlying [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK) API.

## Getting Started
You can quickly start work with the library by following steps listed below.
### Prerequisites
* .NET Core 3.1 or above
### Installing
To install SlideXML, run the following command in the Package Manager Console
```
PM> Install-Package SlideXML
```
### Usage
```C#
// Opens presentation from the file path
using var presentation = new PresentationSL(@"c:\test.pptx");

// Gets the slide collection
var slides = presentation.Slides; 

// Gets number of slides
var numSlides = slides.Count(); 

// Gets the shape collection of the first slide
var shapes = slides[0].Shapes; 

// Prints texts of TextBox shapes on the Debug console
foreach (var sp in shapes)
{
    if (sp.HasTextFrame)
    {
        Debug.WriteLine(sp.TextFrame.Text);
    }
}
```

## Support
* If you have "how-to" questions please post [Stack Overflow](https://stackoverflow.com/) with **slidexml** tag.
* If you get an exception while work with the library's API, then create an issue. You also can send an email message to theadamo86@gmail.com with "SlideXML" subject.

## Author
**Adam Shakhabov** – [adamshakhabov](https://www.linkedin.com/in/adamshakhabov)

## License
[![License](http://img.shields.io/:license-mit-blue.svg?style=flat-square)](http://badges.mit-license.org)