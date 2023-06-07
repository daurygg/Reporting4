using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestingReportSetWithTelerik.Model
{
    public class ModelDate
    {
        DateTime Date1;
        DateTime Date2;

        public DateTime Date11 { get => Date1; set => Date1 = value; }
        public DateTime Date22 { get => Date2; set => Date2 = value; }
    }
}
