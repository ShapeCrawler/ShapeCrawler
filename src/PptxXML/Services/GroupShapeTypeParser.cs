using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using PptxXML.Enums;
using PptxXML.Extensions;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxXML.Services
{
    /// <summary>
    /// Represents a parser of <see cref="P.ShapeTree"/> and <see cref="P.GroupShape"/> instances.
    /// </summary>
    /// <remarks>
    /// <see cref="P.ShapeTree"/> and <see cref="P.GroupShape"/> both derived from <see cref="P.GroupShapeType"/> class.
    /// </remarks>
    public class GroupShapeTypeParser : IGroupShapeTypeParser
    {
        #region Public Methods

        /// <summary>
        /// Creates candidate collection.
        /// </summary>
        /// <param name="groupTypeShape">ShapeTree or GroupShape</param>
        /// <returns></returns>
        public IEnumerable<ElementCandidate> CreateCandidates(P.GroupShapeType groupTypeShape, bool groupParsed = true)
        {
            // Get all element elements
            var allElements = groupTypeShape.Elements<OpenXmlCompositeElement>();

            // Get non-placeholder elements
            var nonPlaceholderElements = allElements.Where(e => e.GetPlaceholderIndex() == null);

            // OLE Objects
            var oleFrames = nonPlaceholderElements.Where(e => e is P.GraphicFrame && e.Descendants<P.OleObject>().Any());
            var oleCandidates = oleFrames.Select(ce => new ElementCandidate
            {
                CompositeElement = ce,
                ElementType = ElementType.OLEObject
            });

            // FILTER PICTURES
            var pictureCandidates = nonPlaceholderElements.Except(oleFrames).Where(e => e is P.Picture || e is P.GraphicFrame && e.Descendants<P.Picture>().Any());
            var graphicFrameImages = pictureCandidates.Where(e => e is P.GraphicFrame).SelectMany(e => e.Descendants<P.Picture>());
            var picAndShapeImages = pictureCandidates.Where(e => e is P.Picture
                                                                 || e is P.Shape && e.Descendants<A.BlipFill>().Any());

            // Picture candidates
            var xmlPictures = graphicFrameImages.Union(picAndShapeImages);
            var picCandidates = xmlPictures.Select(ce => new ElementCandidate
            {
                CompositeElement = ce,
                ElementType = ElementType.Picture
            });

            // Shape candidates
            var xmlShapes = allElements.Except(pictureCandidates)
                                                                     .Where(e => e is P.Shape);
            var shapeCandidates = xmlShapes.Select(ce => new ElementCandidate
            {
                CompositeElement = ce,
                ElementType = ElementType.Shape
            });

            // Table candidates
            var xmlTables = nonPlaceholderElements
                                                            .Where(e => e.GetPlaceholderIndex() == null) // skip placeholders
                                                            .Except(pictureCandidates)
                                                            .Except(xmlShapes)
                                                            .Where(e => e is P.GraphicFrame grFrame && grFrame.Descendants<A.Table>().Any());
            var tableCandidates = xmlTables.Select(ce => new ElementCandidate
            {
                CompositeElement = ce,
                ElementType = ElementType.Table
            });

            // Chart candidates
            var xmlCharts = nonPlaceholderElements
                                                            .Except(pictureCandidates)
                                                            .Except(xmlShapes)
                                                            .Except(xmlTables)
                                                            .Where(e => e is P.GraphicFrame grFrame && grFrame.HasChart());
            var chartCandidates = xmlCharts.Select(ce => new ElementCandidate
            {
                CompositeElement = ce,
                ElementType = ElementType.Chart
            });

            var allCandidates = picCandidates.Union(shapeCandidates).Union(tableCandidates).Union(chartCandidates).Union(oleCandidates);

            // Group candidates
            if (groupParsed)
            {
                var xmlGroupCandidates = nonPlaceholderElements.Where(e => e is P.GroupShape).Select(ce => new ElementCandidate
                {
                    CompositeElement = ce,
                    ElementType = ElementType.Group
                });
                allCandidates = allCandidates.Union(xmlGroupCandidates);
            }

            return allCandidates;
        }

        #endregion Public Methods
    }
}
