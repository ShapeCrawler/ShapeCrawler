using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SlideXML.Models.Settings;
using SlideXML.Services;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideXML.Models.SlideComponents
{
    /// <summary>
    /// Represents a group element.
    /// </summary>
    public class GroupSL
    {
        #region Fields

        private List<ShapeSL> _shapes;

        #region Dependencies

        private readonly IGroupShapeTypeParser _groupShapeTypeParser;
        private readonly IElementFactory _elFactory;
        private readonly IPreSettings _preSettings;
        private readonly OpenXmlCompositeElement _compositeElement;

        #endregion Dependencies

        #endregion Fields

        #region Properties

        /// <summary>
        /// Gets child elements.
        /// </summary>
        public IList<ShapeSL> Shapes
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

        public GroupSL(IGroupShapeTypeParser parser, 
                        IElementFactory elFactory, 
                        OpenXmlCompositeElement compositeElement, 
                        IPreSettings preSettings,
                        SlidePart sldPart)
        {
            _groupShapeTypeParser = parser;
            _elFactory = elFactory;
            _preSettings = preSettings;
            _compositeElement = compositeElement;
        }

        #endregion Constructors

        #region Private Methods

        private void InitChildElements()
        {
            _shapes = new List<ShapeSL>();
            var xmlGroupShape = (P.GroupShape) _compositeElement;
            var tg = xmlGroupShape.GroupShapeProperties.TransformGroup;
            var groupShapeCandidates = _groupShapeTypeParser.CreateCandidates(xmlGroupShape, false); // false is set to avoid parse group in group

            foreach (var ec in groupShapeCandidates)
            {
                ShapeSL newEl = _elFactory.CreateShape(ec, _preSettings);
                newEl.X = newEl.X - tg.ChildOffset.X + tg.Offset.X;
                newEl.Y = newEl.Y - tg.ChildOffset.Y + tg.Offset.Y;
                _shapes.Add(newEl);
            }
        }

        #endregion Private Methods
    }
}
