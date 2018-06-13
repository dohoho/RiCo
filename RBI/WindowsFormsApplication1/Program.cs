using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

using RBI.PRE.subForm;
using RBI.PRE.subForm.InputDataForm;
using DevExpress.XtraSplashScreen;
using System.Data.SqlClient;
using RBI.DAL.MSSQL;
namespace RBI
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                SqlConnection conn = MSSQLDBUtils.GetDBConnection();
                conn.Open();
                conn.Close();
            }
            catch
            {
                MessageBox.Show("Cannot Connect to Database", "Cortek RBI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new test());
            Application.Run(new RibbonForm1());
        }
    }
}
