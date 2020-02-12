using System.Collections.Generic;
using SlideXML.Exceptions;
using SlideXML.Extensions;

namespace SlideXML.Models.TextBody
{
    public class NoTextFrame : ITextFrame
    {
        public IList<Paragraph> Paragraphs => throw new SlideXmlException(ExceptionMessages.NoTextFrame);

        public string Text => throw new SlideXmlException(ExceptionMessages.NoTextFrame);
    }
}