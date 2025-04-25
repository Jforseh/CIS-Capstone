using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BibleCompiler2
{
    public partial class frmSplash : Form
    {
        public frmSplash()
        {
            InitializeComponent();
            this.Text = "Bible Challenge Compiler";
        }
        int closeTime = 20;
        private void btnClose_Click_1_Click(object sender, EventArgs e)
        {
            Close();
        }
        private void frmSplash_Load(object sender, EventArgs e)
        {
            tmrCountDown.Start();
        }

        private void tmrCountDown_Tick(object sender, EventArgs e)
        {
            --closeTime;
            if (closeTime <= 0)
            {
                Close();
            }
        }
    }
}
