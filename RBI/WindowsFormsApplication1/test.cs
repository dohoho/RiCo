using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using RBI.DAL.MSSQL_CAL;
using System.Diagnostics;
using DevExpress.XtraTreeList;
using DevExpress.XtraTreeList.Design;
using DevExpress.XtraTreeList.Menu;
using DevExpress.XtraTreeList.Nodes;
using System.Net;
using RBI.BUS.BUSExcel;
using RBI.Object.ObjectMSSQL;

namespace RBI
{
    public partial class test : Form
    {
        
        public test()
        {
            InitializeComponent();
            float a = 41321.88f;
            int b = (int)a;
            string a1 = b.ToString();
            Console.WriteLine("length " + a1.Length);
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            textBox1.Width += 100;
            Point pt = new Point();
            pt.X = textBox1.Location.X - 100;
            pt.Y = textBox1.Location.Y;
            textBox1.Location = pt;
        }
    }
}
