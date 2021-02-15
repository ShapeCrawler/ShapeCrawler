using System;
using System.Collections.Generic;
using ShapeCrawler.Collections;
using ShapeCrawler.SlideMaster;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Models.Experiment
{
    /// <summary>
    /// Represents a table element on a slide.
    /// </summary>
    public class TableNew : BaseShape
    {
        #region Fields

        private readonly P.GraphicFrame _pGraphicFrame;

        #endregion Fields

        #region Public Properties

        public RowCollection Rows => GetRowsCollection();

        #endregion Public Properties

        internal GroupShapeSc Shape { get; set; }

        #region Constructors

        internal TableNew(P.GraphicFrame pGraphicFrame)
        {
            _pGraphicFrame = pGraphicFrame ?? throw new ArgumentNullException(nameof(pGraphicFrame));
        }

        internal TableNew(SlideMasterSc slideMaster, P.GraphicFrame pGraphicFrame)
        {
            _pGraphicFrame = pGraphicFrame;
        }

        #endregion Constructors

        #region Private Methods

        private RowCollection GetRowsCollection()
        {
            A.Table aTable = _pGraphicFrame.GetFirstChild<A.Graphic>().GraphicData.GetFirstChild<A.Table>();
            IEnumerable<A.TableRow> tableRows = aTable.Elements<A.TableRow>();

            throw new NotImplementedException();
        }

        #endregion Private Methods

        public override long Width => throw new NotImplementedException();

        public override long Height => throw new NotImplementedException();

        public override long X => throw new NotImplementedException();

        public override long Y => throw new NotImplementedException();

        public override GeometryType GeometryType => throw new NotImplementedException();
    }
}