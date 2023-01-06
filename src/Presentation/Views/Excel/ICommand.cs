using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Presentation.Views.Excel
{
    public interface ICommand
    {
        void Execute();

        void Execute(string query);
        DataTable GetCache();
    }
}
