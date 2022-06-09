using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookAutoSig
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
        public void ShowHtml(string html)
        {
            WebBrowser w = new WebBrowser();
            w.Parent = this;
            w.Dock = DockStyle.Fill;
            w.DocumentText = html;
        }
    }
}
