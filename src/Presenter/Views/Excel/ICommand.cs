using System.Data;

namespace Presentation.Views.Excel
{
    public interface ICommand
    {
        void Execute();

        void Execute(string query);
        DataTable GetCache();
    }
}
