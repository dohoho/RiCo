using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Reflection;
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
    public class MyClass
    {
        public string Name { set; get; }
        public static int theValue;
        public void SayHello() { }
    }
    public partial class test : Form
    {
        
        public test()
        {
            InitializeComponent();
            Type type = typeof(MyClass);
            MemberInfo[] members = type.GetMembers();
            Array.ForEach(members, mem =>
            Console.WriteLine(mem.MemberType.ToString().PadRight(12) + ": " + mem));
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            
        }
    }
}
