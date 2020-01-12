using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using PptxXML.Enums;
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
        private readonly IElementFactory _elFactory;
        private SlidePart _sldPart;

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

        private GroupEx(IGroupShapeTypeParser parser, IElementFactory elFactory, OpenXmlCompositeElement compositeElement) 
            : base(ElementType.Group, compositeElement)
        {
            _groupShapeTypeParser = parser;
            _elFactory = elFactory;
        }

        #endregion Constructors

        #region Private Methods

        private void InitChildElements()
        {
            _elements = new List<Element>();
            var xmlGroupShape = (P.GroupShape) CompositeElement;
            var tg = xmlGroupShape.GroupShapeProperties.TransformGroup;
            var groupShapeCandidates = _groupShapeTypeParser.CreateCandidates(xmlGroupShape, false); // false is set to avoid parse group in group

            foreach (var ec in groupShapeCandidates)
            {
                Element newEl = _elFactory.CreateGroupsElement(ec, _sldPart);
                newEl.X = newEl.X - tg.ChildOffset.X + tg.Offset.X;
                newEl.Y = newEl.Y - tg.ChildOffset.Y + tg.Offset.Y;
                _elements.Add(newEl);
            }
        }

        #endregion Private Methods

        #region Builder

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
            public GroupEx Build(P.GroupShape xmlGroupShape, SlidePart sldPart)
            {
                var group = new GroupEx(_parser, _elFactory, xmlGroupShape)
                {
                    _sldPart = sldPart
                };
                var tg = xmlGroupShape.GroupShapeProperties.TransformGroup;
                group.X = tg.Offset.X.Value;
                group.Y = tg.Offset.Y.Value;
                group.Width = tg.Extents.Cx.Value;
                group.Height = tg.Extents.Cy.Value;

                return group;
            }
        }

        #endregion Builder
    }
}
