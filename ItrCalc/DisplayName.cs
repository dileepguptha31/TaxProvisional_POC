using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ItrCalc
{
    [AttributeUsage(AttributeTargets.All/*, AllowMultiple = true*/)]
    public class DisplayNameAttribute : Attribute
    {
        public readonly string DisplayName;

        public DisplayNameAttribute(string displayName)
        {
            this.DisplayName = displayName;
        }
    }
}
