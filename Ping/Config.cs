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
    public partial class Config : Form
    {
        public Config()
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
                //string crypttext = Form1.EncryptString(textBox1.Text, Form1.configPassword);
                XmlDocument doc = new XmlDocument();
                doc.Load("Config.xml");
                //XmlNode node = doc.SelectSingleNode("/appSettings/configuration/ADPassword");
                //node.InnerText = crypttext;
                doc.SelectSingleNode("/appSettings/configuration/DNSServer").InnerText = textBox1.Text;
                doc.SelectSingleNode("/appSettings/configuration/ADServer").InnerText = textBox2.Text;
                doc.SelectSingleNode("/appSettings/configuration/ADDomain").InnerText = textBox3.Text;
                doc.SelectSingleNode("/appSettings/configuration/ADSearch").InnerText = textBox4.Text;
                doc.SelectSingleNode("/appSettings/configuration/PingTimeout").InnerText = Convert.ToString(numericUpDown1.Value);
                doc.SelectSingleNode("/appSettings/configuration/ADsso").InnerText = Convert.ToString(checkBox1.Checked);
                doc.SelectSingleNode("/appSettings/configuration/ADUser").InnerText = textBox5.Text;
                doc.SelectSingleNode("/appSettings/configuration/ADPassword").InnerText = CryptUtils.EncryptString(textBox6.Text, CryptUtils.configPassword);
                doc.Save("Config.xml");
                this.Close();
            }
        }

        private void Config_Load(object sender, EventArgs e)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load("Config.xml");
            textBox1.Text = doc.SelectSingleNode("/appSettings/configuration/DNSServer").InnerText;
            textBox2.Text = doc.SelectSingleNode("/appSettings/configuration/ADServer").InnerText;
            textBox3.Text = doc.SelectSingleNode("/appSettings/configuration/ADDomain").InnerText;
            textBox4.Text = doc.SelectSingleNode("/appSettings/configuration/ADSearch").InnerText;
            numericUpDown1.Value = Convert.ToInt32(doc.SelectSingleNode("/appSettings/configuration/PingTimeout").InnerText);
            checkBox1.Checked = Convert.ToBoolean(doc.SelectSingleNode("/appSettings/configuration/ADsso").InnerText);
            if (checkBox1.Checked==false)
            {
                textBox5.Enabled = true;
                textBox6.Enabled = true;
                textBox5.Text = doc.SelectSingleNode("/appSettings/configuration/ADUser").InnerText;
                textBox6.Text = CryptUtils.DecryptString(doc.SelectSingleNode("/appSettings/configuration/ADPassword").InnerText, CryptUtils.configPassword);
            }
            else
            {
                textBox5.Enabled = false;
                textBox6.Enabled = false;
                textBox5.Text = doc.SelectSingleNode("/appSettings/configuration/ADUser").InnerText;
                textBox6.Text = CryptUtils.DecryptString(doc.SelectSingleNode("/appSettings/configuration/ADPassword").InnerText, CryptUtils.configPassword);
            }
        }

        private void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == false)
            {
                textBox5.Enabled = true;
                textBox6.Enabled = true;
            }
            else
            {
                textBox5.Enabled = false;
                textBox6.Enabled = false;
            }
        }
    }
}
