using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelAdaptor
{
    public class CellReference
    {
        public static CellReference Null = new CellReference();
        public string Column { get; } = "";
        public int Row { get; } = 0;
        public CellReference()
        {
        }
        public CellReference(int row, string column)
        {
            if (IsValidColumn(column) && IsValidRow(row))
            {
                Column = column;
                Row = row;
            }
            else
                throw new ArgumentException("Invalid row or column");
        }
        public CellReference(int row, int column)
        {
            if(IsValidColumn(column) && IsValidRow(row))
            {
                Column = GetColumnName(column);
                Row = row;
            }
            else
                throw new ArgumentException("Invalid row or column");
        }
        public static bool IsValidColumn(string column)
        {
            if(!String.IsNullOrEmpty(column))
            {
                if (column.Length == 1 && char.IsLetter(column[0])) return true;
                if (column.Length == 2 && char.IsLetter(column[0]) && char.IsLetter(column[1])) return true;
            }
            return false;
        }
        public static bool IsValidColumn(int index)
        {
            return index > 0;
        }
        public static bool IsValidRow(int index)
        {
            return index > 0;
        }
        public bool IsNull => Column == "" && Row == 0;
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
            string cn = columnName.ToUpper();
            if (cn == "")
                return 0;
            else if (cn.Length == 1)
                return cn[0] - 'A' + 1;
            else if (columnName.Length == 2)
                return (cn[0] - 'A' + 1) * 26 + cn[1] - 'A' + 1;
            else
                return 0;
        }
        public static string GetColumnName(int column)
        {
            if (column < 1) return "";
            else if (column <= 26) return new string(new char[] { (char)('A' + column - 1) });
            else if (column <= 26 * 26 + 26) return new string(new char[] { (char)('A' + (column - 27) / 26), (char)('A' + (column - 1) % 26) });
            else return "";
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
        public override string ToString()
        {
            if (this == Null)
                return "";
            else
                return Column + Row.ToString();
        }
    }
}
