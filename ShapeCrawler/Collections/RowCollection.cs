using System;
using System.Collections.Generic;
using System.Linq;
using ShapeCrawler.Extensions;
using ShapeCrawler.Tables;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable PossibleMultipleEnumeration

namespace ShapeCrawler.Collections
{
    /// <summary>
    ///     Represents a table rows collection.
    /// </summary>
    public class RowCollection : EditableCollection<RowSc>
    {
        #region Constructors

        internal RowCollection(List<RowSc> rowList)
        {
            CollectionItems = rowList;
        }

        #endregion Constructors

        internal static RowCollection Create(TableSc table, P.GraphicFrame pGraphicFrame)
        {
            IEnumerable<A.TableRow> aTableRows = pGraphicFrame.GetATable().Elements<A.TableRow>();
            var rowList = new List<RowSc>(aTableRows.Count());
            int rowIndex = 0;
            rowList.AddRange(aTableRows.Select(aTblRow => new RowSc(table, aTblRow, rowIndex++)));

            return new RowCollection(rowList);
        }

        #region Public Methods

        /// <summary>
        ///     Removes the specified table row.
        /// </summary>
        /// <param name="row"></param>
        public override void Remove(RowSc row)
        {
            row.ATableRow.Remove();
            CollectionItems.Remove(row);
        }

        /// <summary>
        ///     Removes table row by index.
        /// </summary>
        /// <param name="index"></param>
        public void RemoveAt(int index)
        {
            if (index < 0 || index >= CollectionItems.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            var innerRow = CollectionItems[index];
            Remove(innerRow);
        }

        #endregion Public Methods
    }
}