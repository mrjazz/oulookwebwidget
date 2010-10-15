using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using System.Xml.Serialization;
namespace OutlookPanel
{
    public partial class TaskPanel : UserControl
    {
        public TaskPanel()
        {
            InitializeComponent();
            this.webBrowser.Navigate(Properties.Settings.Default.AddInUrl);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.webBrowser.Navigate(this.textBox1.Text);
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar.Equals((char)13))
            {
                button1_Click(sender, e);
            }
        }



        private void webBrowser1_Navigated(object sender, WebBrowserNavigatedEventArgs e)
        {
            Properties.Settings.Default.AddInUrl = this.webBrowser.Url.ToString();
            this.textBox1.Text = Properties.Settings.Default.AddInUrl;
        }
    }
}
