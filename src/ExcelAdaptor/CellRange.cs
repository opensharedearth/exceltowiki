using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelAdaptor
{
    /// <summary>   The Excel cell range.  A cell range is defined
    ///             by an upper left cell (UL) and a lower right cell (LR).  The
    ///             range can be expressed as a string with the 2 cell references separated by a colon. </summary>
    public class CellRange
    {
        /// <summary>   Gets a reference to the upper left cell. </summary>
        ///
        /// <value> The upper left cell reference. </value>
        public CellReference ULCell { get; } = CellReference.Null;
        /// <summary>   Gets a reference to the lower right cell. </summary>
        ///
        /// <value> The lower right cell reference. </value>
        public CellReference LRCell { get; } = CellReference.Null;
        /// <summary>   Default constructor.  Used to indicate a null range.</summary>
        public CellRange()
        {

        }
        /// <summary>   Constructor. </summary>
        ///
        /// <exception cref="ArgumentNullException">    Thrown when one or more required arguments are
        ///                                             null. </exception>
        /// <exception cref="ArgumentException">        Thrown when one or more arguments have
        ///                                             unsupported or illegal values. </exception>
        ///
        /// <param name="range">    The cell range expressed as a string. The string must contain the upper left and
        ///                         lower right cell references separated by a colon.</param>
        /// <example>
        ///     var cr = new CellRange("A1:C10");
        /// </example>
        public CellRange(string range)
        {
            if (range == null) throw new ArgumentNullException();
            string[] crefs = range.Split(':');
            if (crefs.Length != 2) throw new ArgumentException("Invalid cell range");
            var cell0 = new CellReference(crefs[0]);
            var cell1 = new CellReference(crefs[1]);
            if(cell0 < cell1)
            {
                ULCell = cell0;
                LRCell = cell1;
            }
            else
            {
                ULCell = cell1;
                LRCell = cell0;
            }
        }
        /// <summary>   Determines whether the specified cell range is equal to the current cell range. </summary>
        ///
        /// <param name="obj">  The cell range to compare with the current object. </param>
        ///
        /// <returns>
        /// <see langword="true" /> if the cell ranges are equal; otherwise,
        /// <see langword="false" />.
        /// </returns>
        public override bool Equals(object obj)
        {
            if(obj is CellRange crange)
            {
                return ULCell == crange.ULCell && LRCell == crange.LRCell;
            }
            return base.Equals(obj);
        }
        /// <summary>   Serves as the default hash function. </summary>
        ///
        /// <returns>   A hash code for the cell range. </returns>
        public override int GetHashCode()
        {
            return ULCell.GetHashCode() ^ LRCell.GetHashCode();
        }
        /// <summary>   Returns a string that represents the cell range. </summary>
        ///
        /// <returns>   A cell range expressed as a string </returns>
        public override string ToString()
        {
            return ULCell.ToString() + ":" + LRCell.ToString();
        }
        /// <summary>   Equality operator. </summary>
        ///
        /// <param name="crange0">  The first cell range to compare. </param>
        /// <param name="crange1">  The second cell range to compare. </param>
        ///
        /// <returns>   The result of the operation. </returns>
        static public bool operator==(CellRange crange0, CellRange crange1)
        {
            if ((Object)crange0 == null) return (Object)crange1 == null;
            else return crange0.Equals(crange1);
        }
        /// <summary>   Inequality operator. </summary>
        ///
        /// <param name="crange0">  The first cell range to compare. </param>
        /// <param name="crange1">  The second cell range to compare. </param>
        ///
        /// <returns>   The result of the operation. </returns>
        static public bool operator!=(CellRange crange0, CellRange crange1)
        {
            if ((Object)crange0 == null) return (Object)crange1 != null;
            return !crange0.Equals(crange1);
        }
        /// <summary>   Gets column count of the cell range. </summary>
        ///
        /// <returns>   The column count. </returns>
        public int GetColumnCount()
        {
            return LRCell.ColumnIndex - ULCell.ColumnIndex + 1;
        }
        /// <summary>   Gets row count of the cell range. </summary>
        ///
        /// <returns>   The row count. </returns>
        public int GetRowCount()
        {
            return LRCell.Row - ULCell.Row + 1;
        }
    }
}
