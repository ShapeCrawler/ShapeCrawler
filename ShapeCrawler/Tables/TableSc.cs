using System;
using System.Linq;
using ShapeCrawler.Enums;
using ShapeCrawler.Models;
using ShapeCrawler.Models.SlideComponents;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Tables
{
    /// <summary>
    /// Represents a table element on a slide.
    /// </summary>
    public class TableSc : BaseShape
    {
        #region Fields

        private readonly Lazy<RowCollection> _rowsCollection;
        private readonly P.GraphicFrame _sdkGrFrame;

        #endregion Fields

        #region Properties

        public RowCollection Rows => _rowsCollection.Value;

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initializes an instance of the <see cref="TableSc"/> class.
        /// </summary>
        public TableSc(P.GraphicFrame pGraphicFrame)
        {
            _sdkGrFrame = pGraphicFrame ?? throw new ArgumentNullException(nameof(pGraphicFrame));
            _rowsCollection = new Lazy<RowCollection>(()=>GetRowsCollection());
        }

        #endregion Constructors

        #region Private Methods

        private RowCollection GetRowsCollection()
        {
            var sdkTblRows = _sdkGrFrame.Descendants<A.Table>().First().Elements<A.TableRow>();

            return new RowCollection(sdkTblRows);
        }

        #endregion Private Methods

        public override long Width => throw new NotImplementedException();

        public override long Height => throw new NotImplementedException();

        public override long X => throw new NotImplementedException();

        public override long Y => throw new NotImplementedException();

        public override GeometryType GeometryType => throw new NotImplementedException();
    }
}