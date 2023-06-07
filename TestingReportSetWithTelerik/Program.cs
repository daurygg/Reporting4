using System;
using System.Linq;
using System.Windows.Forms;
using TestingReportSetWithTelerik.Interfaces;

namespace TestingReportSetWithTelerik
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            string SqlconecctionString = "Data Source=GGINC;Initial Catalog=Northwind;Integrated Security=True";
            IRepositorio repository = new Repository(SqlconecctionString);
            RadForm1 Form = new RadForm1(SqlconecctionString, repository);
            Application.Run(Form);
        }

    }
}