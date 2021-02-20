using System;
using System.Collections.Generic;
using System.Text;

namespace exceltowiki
{
    public class CellReference
    {
        public static CellReference Null = new CellReference();
        public string Column { get; } = "";
        public int Row { get; } = 0;
        public CellReference()
        {
        }
        public CellReference(string reference)
        {
            if (reference == null) throw new ArgumentNullException();
            StringBuilder column = new StringBuilder();
            StringBuilder row = new StringBuilder();
            foreach(char c in reference)
            {
                if (char.IsLetter(c)) column.Append(c);
                else if (char.IsDigit(c)) row.Append(c);
            }
            string rawColumn = column.ToString();
            string rawRow = row.ToString();
            if (rawColumn + rawRow != reference || rawColumn.Length < 1 || rawColumn.Length > 2) throw new ArgumentException("Invalid cell reference");
            Column = rawColumn.ToUpper();
            Row = int.Parse(rawRow);
        }
        public override bool Equals(object obj)
        {
            if(obj is CellReference cref)
            {
                return Column == cref.Column && Row == cref.Row;
            }
            return base.Equals(obj);
        }
        public override int GetHashCode()
        {
            return Column.GetHashCode() ^ Row.GetHashCode();
        }
        public static bool operator==(CellReference cref0, CellReference cref1)
        {
            return cref0.Equals(cref1);
        }
        public static bool operator !=(CellReference cref0, CellReference cref1)
        {
            return !cref0.Equals(cref1);
        }
        public static int GetColumnIndex(string columnName)
        {
            if (columnName == "")
                return 0;
            else if (columnName.Length == 1)
                return columnName[0] - 'A' + 1;
            else if (columnName.Length == 2)
                return (columnName[0] - 'A' + 1) * 26 + columnName[1] - 'A' + 1;
            else
                return 0;
        }
        public int ColumnIndex => GetColumnIndex(Column);
        public static bool operator<(CellReference cref0, CellReference cref1)
        {
            return cref0.ColumnIndex < cref1.ColumnIndex || (cref0.Column == cref1.Column && cref0.Row < cref1.Row);
        }
        public static bool operator>(CellReference cref0, CellReference cref1)
        {
            return cref0.ColumnIndex > cref1.ColumnIndex || (cref0.Column == cref1.Column && cref0.Row > cref1.Row);
        }
        public static bool operator<=(CellReference cref0, CellReference cref1)
        {
            return cref0 == cref1 || cref0 < cref1;
        }
        public static bool operator >=(CellReference cref0, CellReference cref1)
        {
            return cref0 == cref1 || cref0 > cref1;
        }
    }
}
