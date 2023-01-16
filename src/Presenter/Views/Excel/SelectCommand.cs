using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Presentation.Views.Excel
{
    public class SelectCommand : CommandBase, ICommand
    {
        DataTable _cache = new DataTable();

        public SelectCommand(string connectionString) : base(connectionString)
        {
        }

        public override void Execute()
        {
            throw new NotImplementedException();
        }

        public override void Execute(string query)
        {
            using(SqlConnection conn = new SqlConnection(base._connectionString))
            {
                conn.Open();
                var cmd = new SqlCommand(query, conn);

                // cmd 객체 SQL 실행, 데이타 Fetch
                SqlDataReader rdr = cmd.ExecuteReader();

                // 그리드에 결과 표시
                DataTable dt = new DataTable();
                dt.Load(rdr);

                _cache = dt;
            }
        }

        public override DataTable GetCache()
        {
            return _cache;
        }
    }
}
