namespace ShapeCrawler.Examples;

public class SlideMasterExamples
{
    [Test, Explicit]
    public void Get_master_shapes()
    {
        using var pres = new Presentation("pres.pptx");
        
        var slideMastersCount = pres.MasterSlides.Count();
        
        var slideMaster = pres.SlideMaster(1);
        
        var masterShapesCount = slideMaster.Shapes.Count;
    }

    [Test, Explicit]
    public void Update_theme()
    {
        using var pres = new Presentation("pres.pptx");
        var theme = pres.SlideMaster(1).Theme;
        
        theme.ColorScheme.Accent1 = "00FF00"; // green
        
        theme.FontScheme.BodyLatinFont = "Calibri";
    }
}