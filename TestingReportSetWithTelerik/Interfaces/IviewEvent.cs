using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestingReportSetWithTelerik.Model;

namespace TestingReportSetWithTelerik.Interfaces
{
    public interface IviewEvent
    {
        DateTime Date1 { get; set; }
        DateTime Date2 { get; set; }
      
    }
}
