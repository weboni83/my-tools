using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Presentation.Views.Excel
{
    public abstract class CommandBase : ICommand
    {
        protected string _connectionString;
        public CommandBase()
        {

        }

        public CommandBase(string connectionString)
        {
            _connectionString = connectionString;
        }

        // Interface 에 위임
        public abstract void Execute();

        public abstract void Execute(string query);

        public abstract DataTable GetCache();
    }
}
