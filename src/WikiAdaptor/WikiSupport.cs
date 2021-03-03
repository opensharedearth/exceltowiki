using System;
using System.Collections.Generic;
using System.Text;

namespace WikiAdaptor
{
    public static class WikiSupport
    {
        public static bool IsValidTitle(string title)
        {
            if(!String.IsNullOrEmpty(title))
            {
                if (title.IndexOfAny("[]{}|#<>%+?".ToCharArray()) < 0) return true;
            }
            return false;
        }
    }
}
