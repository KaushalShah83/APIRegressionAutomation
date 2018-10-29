using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace APIRegressionAutomation
{
    class Program
    {
        [STAThreadAttribute]
        static void Main(string[] args)
        {
            FrmAPITesting frm = new FrmAPITesting();
            frm.ShowDialog();                        
        }
       
    }
}
