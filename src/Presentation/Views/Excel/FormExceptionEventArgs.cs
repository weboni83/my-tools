using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Presentation.Views.Excel
{
    public class FormExceptionEventArgs : EventArgs
    {
        public FormExceptionEventArgs(string message)
        {
            this.Message = message;
        }
        public string Message { get; private set; }
    }
}
