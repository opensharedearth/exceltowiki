using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Text;

namespace exceltowiki
{
    public class SharedString
    {
        private SharedStringItem _item;
        public SharedString()
        {

        }
        public SharedString(SharedStringItem item)
        {
            _item = item;
        }
        public String Text => _item.Text.Text;
        public override string ToString()
        {
            return Text;
        }
    }
}
