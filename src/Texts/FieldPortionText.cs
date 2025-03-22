using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Texts;

internal sealed class FieldPortionText(A.Field aField)
{
    internal string Value
    {
        get
        {
            var aText = aField.GetFirstChild<A.Text>();
        
            return aText == null ? string.Empty : aText.Text;    
        }
    }

    internal void Update(string value)
    {
        aField.GetFirstChild<A.Text>()?.Remove();
        aField.AppendChild(new A.Text { Text = value });
    }
}