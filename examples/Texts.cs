namespace ShapeCrawler.Examples;

public class Texts
{
    [Test, Explicit]
    public void Set_text()
    {
        using var pres = new Presentation("hello world.pptx");
        var slide = pres.Slides.First();
        var shape = slide.Shapes.First();
        
        shape.TextBox!.SetText("A new shape text");
        
        var paragraph = shape.TextBox.Paragraphs[1];
        paragraph.Text = "A new text for second paragraph";
        
        var paragraphPortion = shape.TextBox.Paragraphs.First().Portions.First();
        Console.WriteLine($"Font name: {paragraphPortion.Font!.LatinName}");
        Console.WriteLine($"Font size: {paragraphPortion.Font.Size}"); 
        
        paragraphPortion.Font.IsBold = true;
        
        var fontColor = paragraphPortion.Font.Color.Hex;
        
        paragraphPortion.Font.Color.Update("FF0000");

        pres.Save();
    }

    [Test, Explicit]
    public void Replace_text()
    {
        using var pres = new Presentation("pres.pptx");
        var textBoxes = pres.Slides[0].GetAllTextBoxes();

        foreach (var textFrame in textBoxes)
        {
            textFrame.SetText("some text");
        }

        pres.Save();
    }
}