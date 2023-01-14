using System.IO;
using System.Linq;
using System.Reflection;
using ClosedXML.Excel;
using ShapeCrawler.UnitTests.Helpers;

namespace ShapeCrawler.UnitTests.Helpers;

public abstract class ShapeCrawlerTest
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

    protected static byte[] GetTestBytes(string fileName)
    {
        return GetTestStream(fileName).ToArray();
    }

    protected static MemoryStream GetTestStream(string fileName)
    {
        var assembly = Assembly.GetExecutingAssembly();
        return assembly.GetResourceStream(fileName);
    }

    protected static string GetTestPptxPath(string fileName)
    {
        var assembly = Assembly.GetExecutingAssembly();
        var stream = assembly.GetResourceStream(fileName);
        
        var testPptxPath = Path.GetTempFileName();
        stream.ToFile(testPptxPath);

        return testPptxPath;
    }

    protected static IPresentation SaveAndOpenPresentation(IPresentation presentation)
    {
        var stream = new MemoryStream();
        presentation.SaveAs(stream);

        return SCPresentation.Open(stream);
    }

    private static IPresentation GetPresentationFromAssembly(string fileName)
    {
        var stream = GetTestStream(fileName);

        return SCPresentation.Open(stream);
    }
}

public class ValidationError
{
    public ValidationError(string description, string path)
    {
        this.Description = description;
        this.Path = path;
    }

    public string Path { get; }

    public string Description { get; }
}