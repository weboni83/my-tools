using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Presentation.ViewModels
{
    public class Page
    {
        public virtual int Index { get; set; }
        public virtual int PysicalNumber { get; set; }
        public virtual int LogicalNumber { get; set; }
        public virtual string Chapter { get; set; }
        public virtual string Description { get; set; }
    }
}
