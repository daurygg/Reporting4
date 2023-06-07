using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestingReportSetWithTelerik.Model;

namespace TestingReportSetWithTelerik.Interfaces
{
    public interface IRepositorio
    {

        Task <DataTable> ShowGridview(DateTime date1, DateTime date2);
    }
}
