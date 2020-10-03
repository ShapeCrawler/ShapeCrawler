using System;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;

namespace SlideDotNet.Models.TextBody
{
    public class Bullet
    {
        private readonly A.ParagraphProperties _xmlParagraphProperties;
        private readonly Lazy<string> _colorHex;

        public string ColorHex => _colorHex.Value;

        public Bullet(A.ParagraphProperties xmlParagraphProperties)
        {
            _xmlParagraphProperties = xmlParagraphProperties;
            _colorHex = new Lazy<string>(ParseColorHex);
        }

        private string ParseColorHex()
        {
            var srgbClrElements = _xmlParagraphProperties.Descendants<A.RgbColorModelHex>();
            if (srgbClrElements.Any())
            {
                return srgbClrElements.Single().Val;
            }

            return null;
        }
    }
}