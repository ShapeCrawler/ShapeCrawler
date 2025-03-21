using System;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Paragraphs;

internal sealed class ParagraphLineBreak(A.Break aBreak): IParagraphPortion
{
    public string Text
    {
        get => Environment.NewLine;
        set
        {
            throw new SCException("New Line portion does not support this setter.");
        }
    }

    public ITextPortionFont Font => throw new SCException("New Line portion does not support this property.");

    public IHyperlink? Link => throw new SCException("New Line portion does not support this property.");

    public Color TextHighlightColor
    {
        get => throw new SCException("New Line portion does not support this property.");
        set => throw new SCException("New Line portion does not support this property");
    }

    public void Remove() => aBreak.Remove();
}