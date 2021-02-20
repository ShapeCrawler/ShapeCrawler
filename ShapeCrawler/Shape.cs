using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using ShapeCrawler.Extensions;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Statics;

namespace ShapeCrawler
{
    public abstract class Shape
    {
        protected Shape(OpenXmlCompositeElement pShapeTreeChild)
        {
            PShapeTreeChild = pShapeTreeChild;
        }

        public string CustomData
        {
            get => GetCustomData();
            set => SetCustomData(value);
        }

        public Placeholder Placeholder
        {
            get
            {
                if (PShapeTreeChild.IsPlaceholder())
                {
                    return new Placeholder(PShapeTreeChild);
                }

                return null;
            }
        }

        internal OpenXmlCompositeElement PShapeTreeChild { get; }

        protected void SetCustomData(string value)
        {
            var customDataElement =
                $@"<{ConstantStrings.CustomDataElementName}>{value}</{ConstantStrings.CustomDataElementName}>";
            PShapeTreeChild.InnerXml += customDataElement;
        }

        private string GetCustomData()
        {
            var pattern = @$"<{ConstantStrings.CustomDataElementName}>(.*)<\/{ConstantStrings.CustomDataElementName}>";
            var regex = new Regex(pattern);
            var elementText = regex.Match(PShapeTreeChild.InnerXml).Groups[1];
            if (elementText.Value.Length == 0)
            {
                return null;
            }

            return elementText.Value;
        }
    }
}