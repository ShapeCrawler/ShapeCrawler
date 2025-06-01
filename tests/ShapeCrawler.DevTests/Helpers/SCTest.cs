using System.Reflection;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;
using ShapeCrawler.Presentations;

namespace ShapeCrawler.DevTests.Helpers;

public abstract class SCTest
{
    protected static T GetShape<T>(string presentation, int slideNumber, int shapeId)
    {
        var scPresentation = GetPresentationFromAssembly(presentation);
        var slide = scPresentation.Slides[slideNumber - 1];
        var shape = slide.Shapes.First(sp => sp.Id == shapeId);

        return (T)shape;
    }

    protected static T GetWorksheetCellValue<T>(byte[] workbookByteArray, string cellAddress)
    {
        var stream = new MemoryStream(workbookByteArray);
        var xlWorkbook = new XLWorkbook(stream);
        var cellValue = xlWorkbook.Worksheets.First().Cell(cellAddress).Value;

        return (T)cellValue;
    }

    public static MemoryStream TestAsset(string file)
    {
        var assembly = Assembly.GetExecutingAssembly();
        var stream = assembly.GetResourceStream(file);
        var mStream = new MemoryStream();
        stream!.CopyTo(mStream);
        mStream.Position = 0;

        return mStream;
    }
    
    protected static Presentation TestPresentation(string yamlFile)
    {
        // Read the YAML file content
        var yamlContent = StringOf(yamlFile);
        
        // Create a new presentation
        var presentation = new Presentation();
        
        // Simple parsing of the YAML file - for this specific format
        if (yamlContent.Contains("shapes:"))
        {
            // Parse each shape entry
            var lines = yamlContent.Split('\n');
            int x = 0, y = 0, width = 0, height = 0;
            string text = string.Empty;
            
            for (int i = 0; i < lines.Length; i++)
            {
                var line = lines[i].Trim();
                
                if (line.StartsWith("- ") || (i == lines.Length - 1))
                {
                    // Add previous shape if we have valid dimensions
                    if (i > 0 && width > 0 && height > 0)
                    {
                        presentation.Slide(1).Shapes.AddShape(x, y, width, height, Geometry.Rectangle, text);
                    }
                    
                    // Reset for next shape
                    if (i < lines.Length - 1)
                    {
                        x = y = width = height = 0;
                        text = string.Empty;
                    }
                }
                else if (line.StartsWith("x:"))
                {
                    x = int.Parse(line.Substring(2).Trim());
                }
                else if (line.StartsWith("y:"))
                {
                    y = int.Parse(line.Substring(2).Trim());
                }
                else if (line.StartsWith("width:"))
                {
                    width = int.Parse(line.Substring(6).Trim());
                }
                else if (line.StartsWith("height:"))
                {
                    height = int.Parse(line.Substring(7).Trim());
                }
                else if (line.StartsWith("text:"))
                {
                    text = line.Substring(5).Trim();
                }
            }
            
            // Add the last shape if not added
            if (width > 0 && height > 0)
            {
                presentation.Slide(1).Shapes.AddShape(x, y, width, height, Geometry.Rectangle, text);
            }
        }
        
        return presentation;
    }

    protected static string StringOf(string fileName)
    {
        var stream = TestAsset(fileName);
        return System.Text.Encoding.UTF8.GetString(stream.ToArray());
    }
    
    protected string GetTestPath(string fileName)
    {
        var stream = TestAsset(fileName);
        var path = Path.GetTempFileName();
        File.WriteAllBytes(path, stream.ToArray());

        return path;
    }

    protected static Presentation SaveAndOpenPresentation(IPresentation presentation)
    {
        var stream = new MemoryStream();
        presentation.Save(stream);

        return new Presentation(stream);
    }

    protected static PresentationDocument SaveAndOpenPresentationAsSdk(IPresentation presentation)
    {
        var stream = new MemoryStream();
        presentation.Save(stream);
        stream.Position = 0;

        return PresentationDocument.Open(stream, true);
    }

    private static IPresentation GetPresentationFromAssembly(string fileName)
    {
        var stream = TestAsset(fileName);

        return new Presentation(stream);
    }
}