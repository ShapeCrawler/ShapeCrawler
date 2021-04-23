using System;
using System.Collections.Generic;
using System.Linq;
using ShapeCrawler.Extensions;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable PossibleMultipleEnumeration

namespace ShapeCrawler.Collections
{
    /// <summary>
    ///     Represents a table rows collection.
    /// </summary>
    public class RowCollection : EditableCollection<SCTableRow> // TODO extract interface and convert to internal
    {
        #region Constructors

        internal RowCollection(List<SCTableRow> rowList)
        {
            this.CollectionItems = rowList;
        }

        #endregion Constructors

        internal static RowCollection Create(SlideTable table, P.GraphicFrame pGraphicFrame)
        {
            IEnumerable<A.TableRow> aTableRows = pGraphicFrame.GetATable().Elements<A.TableRow>();
            var rowList = new List<SCTableRow>(aTableRows.Count());
            int rowIndex = 0;
            rowList.AddRange(aTableRows.Select(aTblRow => new SCTableRow(table, aTblRow, rowIndex++)));

            return new RowCollection(rowList);
        }

        #region Public Methods

        /// <inheritdoc/>
        public override void Remove(SCTableRow scTableRow)
        {
            scTableRow.ATableRow.Remove();
            this.CollectionItems.Remove(scTableRow);
        }

        /// <summary>
        ///     Removes table row by index.
        /// </summary>
        public void RemoveAt(int index)
        {
            if (index < 0 || index >= this.CollectionItems.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            var innerRow = this.CollectionItems[index];
            this.Remove(innerRow);
        }

        #endregion Public Methods
    }
}