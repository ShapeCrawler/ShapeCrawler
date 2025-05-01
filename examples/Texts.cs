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
        var textBoxes = pres.Slides[0].GetTextBoxes();

        foreach (var textFrame in textBoxes)
        {
            textFrame.SetText("some text");
        }

        pres.Save();
    }

    [Test, Explicit]
    public void Get_text_margins()
    {
        using var pres = new Presentation("some.pptx");
        var shape = pres.Slide(1).Shape("AutoShape 1");
        var textFrame = shape.TextBox!;

        var leftMargin = textFrame.LeftMargin;
        var topMargin = textFrame.TopMargin;
    }

    [Test, Explicit]
    public void Set_autofit()
    {
        using var pres = new Presentation("some.pptx");
        var textBox = pres.Slide(1).Shapes.Shape("AutoShape 1").TextBox!;
        
        textBox.AutofitType = AutofitType.Resize;
    }

    [Test, Explicit]
    public void Get_alignment()
    {
        using var pres = new Presentation("text.pptx");
        var shape = pres.Slide(1).Shapes.Shape("TextBox 1");
        var paragraph = shape.TextBox!.Paragraphs[0];
        
        var alignment = paragraph.HorizontalAlignment;
        
        paragraph.HorizontalAlignment = TextHorizontalAlignment.Center;
    }
    
    [Test, Explicit]
    public void Set_hyperlink()
    {
        using var pres = new Presentation("text.pptx");
        var shape = pres.Slide(1).Shapes.Shape("TextBox 1");
        var paragraph = shape.TextBox!.Paragraphs[0];
        
        paragraph.Portions[0].Link!.AddFile("https://github.com/ShapeCrawler/ShapeCrawler");
    }
    
    [Test, Explicit]
    public void Set_paragraph_bullet()
    {
        using var pres = new Presentation("text.pptx");
        var textBox = pres.Slide(1).Shapes.Shape("TextBox 1");
        var bullet = textBox.TextBox!.Paragraphs[0].Bullet;

        bullet.Type = BulletType.Character;
        bullet.Character = "*";
        bullet.Size = 100;
        bullet.FontName = "Arial";
    }

    [Test, Explicit]
    public void Set_text_highlight()
    {
        using var pres = new Presentation("some.pptx");
        var textBox = pres.Slide(1).Shapes.Shape("TextBox 1");
        var textPortion = textBox.TextBox!.Paragraphs[0].Portions[0];

        // Set predefined color
        textPortion.TextHighlightColor = Color.Black;

        // Set color by its hexadecimal code
        var greenColor = Color.FromHex("00ff00");
        textPortion.TextHighlightColor = greenColor;
    }
}