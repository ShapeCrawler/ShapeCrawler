using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using ShapeCrawler.Extensions;

namespace ShapeCrawler.Tests.Helpers;

public abstract class ShapeCrawlerTest
{
    protected static T GetShape<T>(string presentation, int slideNumber, int shapeId)
    {
        var scPresentation = GetPresentationFromAssembly(presentation);
        var slide = scPresentation.Slides[slideNumber - 1];
        var shape = slide.Shapes.First(sp => sp.Id == shapeId);

        return (T) shape;
    }
        
    protected static IAutoShape GetAutoShape(string presentation, int slideNumber, int shapeId)
    {
        var scPresentation = GetPresentationFromAssembly(presentation);
        var slide = scPresentation.Slides.First(s => s.Number == slideNumber);
        var shape = slide.Shapes.First(sp => sp.Id == shapeId);

        return (IAutoShape) shape;
    }

    protected static T GetWorksheetCellValue<T>(byte[] workbookByteArray, string cellAddress)
    {
        var stream = new MemoryStream(workbookByteArray);
        var xlWorkbook = new XLWorkbook(stream);
        var cellValue = xlWorkbook.Worksheets.First().Cell(cellAddress).Value;

        return (T)cellValue;
    }

    protected static byte[] GetTestBytes(string fileName)
    {
        return GetTestStream(fileName).ToArray();
    }
        
    protected static MemoryStream GetTestStream(string fileName)
    {
        var assembly = Assembly.GetExecutingAssembly();
        var path = assembly.GetManifestResourceNames().First(r => r.EndsWith(fileName, StringComparison.Ordinal));
        var stream = assembly.GetManifestResourceStream(path);
        var mStream = new MemoryStream();
        stream!.CopyTo(mStream);
        mStream.Position = 0;

        return mStream;
    }
        
    protected static string GetTestPptxPath(string fileName)
    {
        var assembly = Assembly.GetExecutingAssembly();
        var path = assembly.GetManifestResourceNames().First(r => r.EndsWith(fileName, StringComparison.Ordinal));
        var stream = assembly.GetManifestResourceStream(path);
        var testPptxPath = Path.GetTempFileName();
        stream.SaveToFile(testPptxPath);

        return testPptxPath;
    }
        
    protected static IPresentation SaveAndOpenPresentation(IPresentation presentation)
    {
        var stream = new MemoryStream();
        presentation.SaveAs(stream);
            
        return SCPresentation.Open(stream);
    }

    protected static List<ValidationErrorInfo> Validate(IPresentation pres)
    {
        var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
        var errors = validator.Validate(pres.SDKPresentation);

        return errors.ToList();
    }
    
    private static IPresentation GetPresentationFromAssembly(string fileName)
    {
        var stream = GetTestStream(fileName);
            
        return SCPresentation.Open(stream);
    }
    
    
}