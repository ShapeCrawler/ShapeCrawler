using System.Collections.Generic;
using SlideXML.Exceptions;
using SlideXML.Extensions;

namespace SlideXML.Models.TextBody
{
    public class NoTextFrame : ITextFrame
    {
        public IList<ParagraphSL> Paragraphs => throw new SlideXMLException(ExceptionMessages.NoTextFrame);

        public string Text => throw new SlideXMLException(ExceptionMessages.NoTextFrame);
    }
}