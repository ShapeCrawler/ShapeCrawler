using System.Collections.Generic;
using System.Linq;
using SlideXML.Enums;
using SlideXML.Models.Settings;
using SlideXML.Models.TableComponents;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideXML.Models.Elements
{
    /// <summary>
    /// Represents a table element on a slide.
    /// </summary>
    public class TableEx: Element
    {
        #region Fields

        private List<RowEx> _rows;

        #endregion Fields

        #region Properties

        public IList<RowEx> Rows
        {
            get
            {
                if (_rows == null)
                {
                    ParseRows();
                }

                return _rows;
            }
        }

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initialise an instance of <see cref="TableEx"/> class.
        /// </summary>
        public TableEx(P.GraphicFrame xmlGrFrame, ElementSettings elSettings) : base(ElementType.Table, xmlGrFrame, elSettings) { }

        #endregion Constructors

        #region Private Methods

        private void ParseRows()
        {
            var xmlRows = CompositeElement.Descendants<A.Table>().Single().Elements<A.TableRow>();
            _rows = new List<RowEx>(xmlRows.Count());
            foreach (var r in xmlRows)
            {
                _rows.Add(new RowEx(r, ElementSettings));
            }
        }

        #endregion Private Methods
    }
}