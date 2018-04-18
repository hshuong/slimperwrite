using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace slimperwrite
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (txtBoxSymbolsList.TextLength > 0)
            {
                // home
                //Program.RunSlimpertowritefile(@"I:\web load\slimerjs-1.0.0\", txtBoxSymbolsList.Text);
                // office
                Program.RunSlimpertowritefile(@"G:\webload\slimerjs-1.0.0\", txtBoxSymbolsList.Text);
                // home
                //Program.CreateRowNameInStatementRowTable(@"I:\web load\slimerjs-1.0.0\");
                // office
                //Program.CreateRowNameInStatementRowTable(@"G:\webload\slimerjs-1.0.0\");
            }
            else
            {
                MessageBox.Show("Chưa có CK để tải báo cáo tài chính");
            }
        }

    }
}
