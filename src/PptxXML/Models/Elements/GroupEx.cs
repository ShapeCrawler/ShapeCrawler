using System.Collections.Generic;
using PptxXML.Enums;
using PptxXML.Exceptions;
using PptxXML.Services;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxXML.Models.Elements
{
    /// <summary>
    /// Represents a group element.
    /// </summary>
    public class GroupEx: Element
    {
        #region Fields

        private List<Element> _elements;

        #region Dependencies

        private readonly IGroupShapeTypeParser _groupShapeTypeParser;
        private readonly IElementFactory _elCreator;

        #endregion Dependencies

        #endregion Fields

        #region Properties

        /// <summary>
        /// Gets child elements.
        /// </summary>
        public IList<Element> Elements
        {
            get
            {
                if (_elements == null)
                {
                    InitChildElements();
                }

                return _elements;
            }
        }

        #endregion Properties

        #region Constructors

        private GroupEx(IGroupShapeTypeParser parser, IElementFactory elFactory) : base(ElementType.Group)
        {
            _groupShapeTypeParser = parser;
            _elCreator = elFactory;
        }

        #endregion Constructors

        /// <summary>
        /// Represents a builder of the <see cref="GroupEx"/> class.
        /// </summary>
        /// <returns>A new instance of the <see cref="GroupEx"/> class.</returns>
        public class Builder : IGroupBuilder
        {
            private readonly IGroupShapeTypeParser _parser;
            private readonly IElementFactory _elFactory;

            public Builder(IGroupShapeTypeParser parser, IElementFactory elFactory)
            {
                _parser = parser;
                _elFactory = elFactory;
            }

            /// <summary>
            /// Builds a new instance of the <see cref="GroupEx"/> class.
            /// </summary>
            /// <returns></returns>
            public GroupEx Build(P.GroupShape xmlGroupShape)
            {
                var group = new GroupEx(_parser, _elFactory)
                {
                    XmlCompositeElement = xmlGroupShape
                };
                var tg = xmlGroupShape.GroupShapeProperties.TransformGroup;
                group.X = tg.Offset.X.Value;
                group.Y = tg.Offset.Y.Value;
                group.Width = tg.Extents.Cx.Value;
                group.Height = tg.Extents.Cy.Value;

                return group;
            }
        }

        #region Private Methods

        private void InitChildElements()
        {
            _elements = new List<Element>();
            var xmlGroupShape = (P.GroupShape) XmlCompositeElement;
            var tg = xmlGroupShape.GroupShapeProperties.TransformGroup;
            var groupShapeCandidates = _groupShapeTypeParser.CreateCandidates(xmlGroupShape, false); // false is set to avoid parse group in group

            // TODO: delete this copy/past
            foreach (var c in groupShapeCandidates)
            {
                Element el;
                switch (c.ElementType)
                {
                    case ElementType.Shape:
                    {
                        el = _elCreator.CreateShape(c);
                        break;
                    }
                    case ElementType.Chart:
                    {
                        el = _elCreator.CreateChart(c);
                        break;
                    }
                    case ElementType.Table:
                    {
                        el = _elCreator.CreateTable(c);
                        break;
                    }
                    case ElementType.Picture:
                    {
                        el = _elCreator.CreatePicture(c);
                        break;
                    }
                    default:
                        throw new PptxXMLException(nameof(ElementType));
                }

                //TODO: parsed x,y,w,h values for child elements (https://github.com/adamshakhabov/PptxXML/issues/7)

                _elements.Add(el);
            }
        }

        #endregion Private Methods
    }
}
