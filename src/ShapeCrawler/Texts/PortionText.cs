using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Texts;

internal class PortionText
{
    private readonly A.Field aField;

    internal PortionText(A.Field aField)
    {
        this.aField = aField;
    }
    
    internal string Text()
    {
        var aText = this.aField.GetFirstChild<A.Text>();
        
        return aText == null ? string.Empty : aText.Text;
    }

    internal void Update(string newText)
    {
        this.aField.GetFirstChild<A.Text>()?.Remove();
        this.aField.AppendChild(new A.Text { Text = newText });
    }
}