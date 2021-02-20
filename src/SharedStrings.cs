using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Text;

namespace exceltowiki
{
    public class SharedStrings : List<SharedString>
    {
        public SharedStrings()
        {

        }
        public SharedStrings(SharedStringTable sst)
        {
            foreach(SharedStringItem item in sst.ChildElements.OfType<SharedStringItem>())
            {
                Add(new SharedString(item));
            }
        }
        public SharedString this[string index]
        { 
            get
            {
                if(int.TryParse(index, out int i))
                {
                    return this[i];
                }
                return new SharedString();
            }
        }
    }
}
