using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Data.Common;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using TestingReportSetWithTelerik.Model;
using DataTable = System.Data.DataTable;
using TestingReportSetWithTelerik.Interfaces;
namespace TestingReportSetWithTelerik
{
    public abstract class RepositoryBase
    {
        protected string connectionstring;
    }

    public class Repository : RepositoryBase, IRepositorio
    {
        DataTable datatable = new DataTable();

        public Repository(string connectionString)
        {
            this.connectionstring = connectionString;
        }


        public async Task <DataTable> ShowGridview(DateTime date1, DateTime date2)
        {
            string StoreprocedureName = "ProductsSold";
            SqlConnection connection = new SqlConnection(connectionstring);
           await connection.OpenAsync();
            SqlCommand command = new SqlCommand(StoreprocedureName, connection);
            command.CommandType = CommandType.StoredProcedure;
            command.Parameters.Add("@date1", SqlDbType.DateTime).Value = date1;//"1990-04-01";
            command.Parameters.Add("@date2", SqlDbType.DateTime).Value = date2;//"2000-04-01";
            SqlDataAdapter dataAdapter = new SqlDataAdapter(command);
            
                dataAdapter.Fill(datatable);
            
           

            return datatable;
        }
    }
}
