using System.Collections.Generic;
using DocumentFormat.OpenXml;
using SlideDotNet.Models.Settings;
using SlideDotNet.Services;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideDotNet.Models.SlideComponents
{
    /// <summary>
    /// Represents a group element.
    /// </summary>
    public class Group
    {
        #region Fields

        private List<Shape> _shapes;

        #region Dependencies

        private readonly IXmlGroupShapeTypeParser _groupShapeTypeParser;
        private readonly IElementFactory _elFactory;
        private readonly IParents _parents;
        private readonly OpenXmlCompositeElement _compositeElement;

        #endregion Dependencies

        #endregion Fields

        #region Properties

        /// <summary>
        /// Gets child elements.
        /// </summary>
        public IList<Shape> Shapes
        {
            get
            {
                if (_shapes == null)
                {
                    InitChildElements();
                }

                return _shapes;
            }
        }

        #endregion Properties

        #region Constructors

        public Group(IXmlGroupShapeTypeParser parser, 
                        IElementFactory elFactory, 
                        OpenXmlCompositeElement compositeElement, 
                        IParents parents)
        {
            _groupShapeTypeParser = parser;
            _elFactory = elFactory;
            _parents = parents;
            _compositeElement = compositeElement;
        }

        #endregion Constructors

        #region Private Methods

        private void InitChildElements()
        {
            _shapes = new List<Shape>();
            var xmlGroupShape = (P.GroupShape) _compositeElement;
            var tg = xmlGroupShape.GroupShapeProperties.TransformGroup;
            var groupShapeCandidates = _groupShapeTypeParser.CreateElementCandidates(xmlGroupShape, false); // false is set to avoid parse group in group

            foreach (var ec in groupShapeCandidates)
            {
                Shape newEl = _elFactory.ElementFromCandidate(ec, _parents);
                newEl.X = newEl.X - tg.ChildOffset.X + tg.Offset.X;
                newEl.Y = newEl.Y - tg.ChildOffset.Y + tg.Offset.Y;
                _shapes.Add(newEl);
            }
        }

        #endregion Private Methods
    }
}
