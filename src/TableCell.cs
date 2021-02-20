using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Text;

namespace exceltowiki
{
    public class TableCell
    {
        public static TableCell Null = new TableCell();
        CellReference _reference;
        Cell _cell = null;
        CellValues _dataType = CellValues.Error;
        CellValue _value = null;
        SharedStrings _sst = null;
        public TableCell()
        {

        }
        public TableCell(Cell cell, SharedStrings sst = null)
        {
            _cell = cell;
            _reference = new CellReference(cell.CellReference.Value);
            if((object)cell.DataType == null)
            {
                _dataType = CellValues.Number;
            }
            else
            {
                _dataType = cell.DataType;
            }
            _value = cell.CellValue;
            _sst = sst;
        }
        public object Value
        {
            get
            {
                if(_value != null)
                {
                    switch (_dataType)
                    {
                        case CellValues.SharedString:
                            return _sst != null ? _sst[_value.Text] : (object)"";
                        default:
                            return _value.Text;
                    }
                }
                else
                {
                    return "";
                }
            }
        }
        public int RowIndex => _reference.Row;
        public int ColumnIndex => _reference.ColumnIndex;
        public string Column => _reference.Column;
        public override string ToString()
        {
            return Value.ToString();
        }
    }
}
