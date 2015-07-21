using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Ianchi.WebDialerAddIn
{
    public partial class frmDialerOptions : Form
    {
        public frmDialerOptions()
        {
            InitializeComponent();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnAccept_Click(object sender, EventArgs e)
        {
            this.webDialerOptions.Apply();
            this.Close();
        }
    }
}
