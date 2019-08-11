using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace Pinger
{
    public partial class Password : Form
    {
        public Password()
        {
            InitializeComponent();
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Yes or no", "Save and exit?", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == System.Windows.Forms.DialogResult.Yes)
            {
                string crypttext = Form1.EncryptString(textBox1.Text, Form1.configPassword);
                XmlDocument doc = new XmlDocument();
                doc.Load("Config.xml");
                XmlNode node = doc.SelectSingleNode("/appSettings/configuration/ADPassword");
                node.InnerText = crypttext;
                doc.Save("Config.xml");
                this.Close();
            }
        }
    }
}
