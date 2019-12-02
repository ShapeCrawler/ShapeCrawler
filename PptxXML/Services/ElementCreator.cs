using System.Linq;
using DocumentFormat.OpenXml;
using objectEx.Extensions;
using PptxXML.Entities.Elements;
using PptxXML.Exceptions;
using PptxXML.Extensions;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxXML.Services
{
    /// <summary>
    /// Represent <see cref="Element"/> instance creator.
    /// </summary>
    public class ElementCreator
    {
        /// <summary>
        /// Create <see cref="Element"/> instance.
        /// </summary>
        /// <param name="xmlCompositeElement"></param>
        /// <returns></returns>
        public Element Create(OpenXmlCompositeElement xmlCompositeElement)
        {
            xmlCompositeElement.ThrowIfNull(nameof(xmlCompositeElement));

            Element element;
            switch (xmlCompositeElement)
            {
                // Group
                case P.GroupShape _:
                    element = new Group(xmlCompositeElement);
                    return element;

                // Shape
                case P.Shape xmlShape:
                    element = new Shape(xmlShape);
                    return element;
            }
            // Chart
            if (xmlCompositeElement.IsChart())
            {
                element = new Chart(xmlCompositeElement);
                return element;
            }
            // Table
            if (xmlCompositeElement.IsTable())
            {
                element = new Table(xmlCompositeElement);
                return element;
            }
            // Picture
            if (xmlCompositeElement is P.Picture
                || xmlCompositeElement is P.GraphicFrame && xmlCompositeElement.Descendants<P.Picture>().Any())
            {
                element = new Picture(xmlCompositeElement);
                return element;
            }

            throw new TypeException();
        }
    }
}
