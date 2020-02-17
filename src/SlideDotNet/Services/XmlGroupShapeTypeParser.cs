using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using SlideDotNet.Enums;
using SlideDotNet.Extensions;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
// ReSharper disable PossibleMultipleEnumeration

namespace SlideDotNet.Services
{
    /// <summary>
    /// Represents a parser of <see cref="P.ShapeTree"/> and <see cref="P.GroupShape"/> instances.
    /// </summary>
    /// <remarks>
    /// <see cref="P.ShapeTree"/> and <see cref="P.GroupShape"/> both derived from <see cref="P.GroupShapeType"/> class.
    /// </remarks>
    public class XmlGroupShapeTypeParser : IXmlGroupShapeTypeParser
    {
        #region Public Methods

        /// <summary>
        /// Creates the collection of instance of the <see cref="ElementCandidate"/> class.
        /// </summary>
        /// <param name="xmlGroupTypeShape">Instance of <see cref="P.ShapeTree"/> or <see cref="P.GroupShape"/> class.</param>
        /// <param name="groupParsed"></param>
        /// <returns></returns>
        public IEnumerable<ElementCandidate> CreateElementCandidates(P.GroupShapeType xmlGroupTypeShape, bool groupParsed = true)
        {
            var allXmlElements = xmlGroupTypeShape.Elements<OpenXmlCompositeElement>();
            var xmlGraphicFrameElements = allXmlElements.OfType<P.GraphicFrame>();

            // OLE objects
            var xmlOleGraphicFrames = xmlGraphicFrameElements.Where(e => e.Descendants<P.OleObject>().Any());
            var oleCandidates = xmlOleGraphicFrames.Select(graphicFrame => new ElementCandidate
            {
                XmlElement = graphicFrame,
                ElementType = ElementType.OLEObject
            });

            // Pictures
            var picGraphicFrames = xmlGraphicFrameElements.Except(xmlOleGraphicFrames).SelectMany(e => e.Descendants<P.Picture>());
            var allXmlPicElements = allXmlElements.OfType<P.Picture>().Union(picGraphicFrames);
            var picCandidates = allXmlPicElements.Select(xmlPic => new ElementCandidate
            {
                XmlElement = xmlPic,
                ElementType = ElementType.Picture
            });

            // AutoShapes
            var xmlShapes = allXmlElements.OfType<P.Shape>();
            var autoShapeCandidates = xmlShapes.Select(xmlShape => new ElementCandidate
            {
                XmlElement = xmlShape,
                ElementType = ElementType.AutoShape
            });

            // Tables
            var xmlTables = xmlGraphicFrameElements.Where(g => g.Descendants<A.Table>().Any());
            var tableCandidates = xmlTables.Select(graphicFrame => new ElementCandidate
            {
                XmlElement = graphicFrame,
                ElementType = ElementType.Table
            });

            // Charts
            var xmlCharts = xmlGraphicFrameElements.Where(g => g.HasChart());
            var chartCandidates = xmlCharts.Select(graphicFrame => new ElementCandidate
            {
                XmlElement = graphicFrame,
                ElementType = ElementType.Chart
            });

            var allCandidates = picCandidates.Union(autoShapeCandidates).Union(tableCandidates).Union(chartCandidates).Union(oleCandidates);

            // Groups
            if (groupParsed)
            {
                var xmlGroupCandidates = allXmlElements.OfType<P.GroupShape>().Select(groupShape => new ElementCandidate
                {
                    XmlElement = groupShape,
                    ElementType = ElementType.Group
                });
                allCandidates = allCandidates.Union(xmlGroupCandidates);
            }

            return allCandidates;
        }

        #endregion Public Methods
    }
}
