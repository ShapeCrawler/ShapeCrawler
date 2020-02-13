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
    public class Group
    {
        #region Fields

        private List<SlideElement> _shapes;

        #region Dependencies

        private readonly IXmlGroupShapeTypeParser _groupShapeTypeParser;
        private readonly IElementFactory _elFactory;
        private readonly IPreSettings _preSettings;
        private readonly OpenXmlCompositeElement _compositeElement;

        #endregion Dependencies

        #endregion Fields

        #region Properties

        /// <summary>
        /// Gets child elements.
        /// </summary>
        public IList<SlideElement> Shapes
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
            _shapes = new List<SlideElement>();
            var xmlGroupShape = (P.GroupShape) _compositeElement;
            var tg = xmlGroupShape.GroupShapeProperties.TransformGroup;
            var groupShapeCandidates = _groupShapeTypeParser.CreateCandidates(xmlGroupShape, false); // false is set to avoid parse group in group

            foreach (var ec in groupShapeCandidates)
            {
                SlideElement newEl = _elFactory.CreateShape(ec, _preSettings);
                newEl.X = newEl.X - tg.ChildOffset.X + tg.Offset.X;
                newEl.Y = newEl.Y - tg.ChildOffset.Y + tg.Offset.Y;
                _shapes.Add(newEl);
            }
        }

        #endregion Private Methods
    }
}
