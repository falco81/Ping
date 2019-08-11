﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.DirectoryServices;
using System.Net;
using System.Net.NetworkInformation;
using System.Xml;
using System.Threading;
using System.IO;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using DnsClient;
using System.Security.Cryptography;

namespace Pinger
{


    public partial class Form1 : Form
    {

        private string Excel03ConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
        private string Excel07ConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0;HDR={1}'";
        public static string configPassword = "muGcqhykTkOHM2pD7UXaBWXU";
        private static byte[] _salt = Encoding.ASCII.GetBytes("0123456789abcdef");

        public void ClearGrid()
        {
           /* int clrowCount = dataGridView1.Rows.Count;
            for (int cln = 0; cln < clrowCount; cln++)
            {
                if (dataGridView1.Rows[0].IsNewRow == false)
                    dataGridView1.Rows.RemoveAt(0);
            }*/
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.ColumnCount = 0;
            //dataGridView1.Refresh();
           
        }

        public void RefreshOU()
        {
            comboBox1.Items.Clear();
            XmlDocument doc = new XmlDocument();
            doc.Load("Config.xml");
            string ADServer = doc.SelectSingleNode("/appSettings/configuration/ADServer").InnerText;
            string ADSearch = doc.SelectSingleNode("/appSettings/configuration/ADSearch").InnerText;
            string ADUser = doc.SelectSingleNode("/appSettings/configuration/ADUser").InnerText;
            //string ADPassword = doc.SelectSingleNode("/appSettings/configuration/ADPassword").InnerText;
            string ADPassword = DecryptString(doc.SelectSingleNode("/appSettings/configuration/ADPassword").InnerText, configPassword);
            bool ADSsso = Convert.ToBoolean(doc.SelectSingleNode("/appSettings/configuration/ADsso").InnerText);
            if (ADSsso)
            {
                DirectoryEntry entry = new DirectoryEntry("LDAP://" + ADServer + "/" + ADSearch);
                DirectorySearcher mySearcher = new DirectorySearcher(entry);
                mySearcher.Filter = ("(objectClass=organizationalUnit)");
                mySearcher.SizeLimit = int.MaxValue;
                mySearcher.PageSize = int.MaxValue;

                foreach (SearchResult resEnt in mySearcher.FindAll())
                {
                    string OUName = resEnt.GetDirectoryEntry().Name;
                    comboBox1.Items.Add(OUName.Remove(0, 3));
                }

                mySearcher.Dispose();
                entry.Dispose();
                comboBox1.SelectedIndex = 0;
            }
            else
            {
                DirectoryEntry entry = new DirectoryEntry("LDAP://" + ADServer + "/" + ADSearch, ADUser, ADPassword);
                DirectorySearcher mySearcher = new DirectorySearcher(entry);
                mySearcher.Filter = ("(objectClass=organizationalUnit)");
                mySearcher.SizeLimit = int.MaxValue;
                mySearcher.PageSize = int.MaxValue;

                foreach (SearchResult resEnt in mySearcher.FindAll())
                {
                    string OUName = resEnt.GetDirectoryEntry().Name;
                    comboBox1.Items.Add(OUName.Remove(0, 3));
                }

                mySearcher.Dispose();
                entry.Dispose();
                comboBox1.SelectedIndex = 0;
            }
        }

        public void ReadPC()
        {
            ClearGrid();
            XmlDocument doc = new XmlDocument();
            doc.Load("Config.xml");
            string ADServer = doc.SelectSingleNode("/appSettings/configuration/ADServer").InnerText;
            string ADSearch = doc.SelectSingleNode("/appSettings/configuration/ADSearch").InnerText;
            string ADUser = doc.SelectSingleNode("/appSettings/configuration/ADUser").InnerText;
            //string ADPassword = doc.SelectSingleNode("/appSettings/configuration/ADPassword").InnerText;
            string ADPassword = DecryptString(doc.SelectSingleNode("/appSettings/configuration/ADPassword").InnerText, configPassword);
            bool ADSsso = Convert.ToBoolean(doc.SelectSingleNode("/appSettings/configuration/ADsso").InnerText);
            if (ADSsso)
            {
                DirectoryEntry entry = new DirectoryEntry("LDAP://" + ADServer + "/OU=" + comboBox1.Text + "," + ADSearch);
                DirectorySearcher mySearcher = new DirectorySearcher(entry);
                mySearcher.Filter = ("(objectCategory=Computer)");
                mySearcher.SizeLimit = int.MaxValue;
                mySearcher.PageSize = int.MaxValue;
                int index = 0;
                dataGridView1.ColumnCount = 3;
                dataGridView1.ColumnHeadersVisible = true;



                // Set the column header style.
                DataGridViewCellStyle columnHeaderStyle = new DataGridViewCellStyle();

                columnHeaderStyle.BackColor = Color.Beige;
                columnHeaderStyle.Font = new Font("Verdana", 10, FontStyle.Bold);
                dataGridView1.ColumnHeadersDefaultCellStyle = columnHeaderStyle;

                // Set the column header names.
                dataGridView1.Columns[0].Name = "Inv. č.";
                dataGridView1.Columns[1].Name = "IP";
                dataGridView1.Columns[2].Name = "Ping";

                try
                {
                    SearchResultCollection results = mySearcher.FindAll();
                    {



                        foreach (SearchResult resEnt in mySearcher.FindAll())
                        {
                            string PCName = resEnt.GetDirectoryEntry().Name;


                            dataGridView1.Rows.Add();
                            dataGridView1.Rows[index].Cells["Inv. č."].Value = PCName.Remove(0, 3);
                            index++;
                        }

                        mySearcher.Dispose();
                        entry.Dispose();
                    }
                }
                catch
                {
                    dataGridView1.Rows.Add();
                    dataGridView1.Rows[index].Cells["Inv. č."].Value = "Nenalezeno žádné PC";
                }

            }
            else
            {
                DirectoryEntry entry = new DirectoryEntry("LDAP://" + ADServer + "/OU=" + comboBox1.Text + "," + ADSearch, ADUser, ADPassword);
                DirectorySearcher mySearcher = new DirectorySearcher(entry);
                mySearcher.Filter = ("(objectCategory=Computer)");
                mySearcher.SizeLimit = int.MaxValue;
                mySearcher.PageSize = int.MaxValue;
                int index = 0;
                dataGridView1.ColumnCount = 3;
                dataGridView1.ColumnHeadersVisible = true;



                // Set the column header style.
                DataGridViewCellStyle columnHeaderStyle = new DataGridViewCellStyle();

                columnHeaderStyle.BackColor = Color.Beige;
                columnHeaderStyle.Font = new Font("Verdana", 10, FontStyle.Bold);
                dataGridView1.ColumnHeadersDefaultCellStyle = columnHeaderStyle;

                // Set the column header names.
                dataGridView1.Columns[0].Name = "Inv. č.";
                dataGridView1.Columns[1].Name = "IP";
                dataGridView1.Columns[2].Name = "Ping";

                try
                {
                    SearchResultCollection results = mySearcher.FindAll();
                    {



                        foreach (SearchResult resEnt in mySearcher.FindAll())
                        {
                            string PCName = resEnt.GetDirectoryEntry().Name;


                            dataGridView1.Rows.Add();
                            dataGridView1.Rows[index].Cells["Inv. č."].Value = PCName.Remove(0, 3);
                            index++;
                        }

                        mySearcher.Dispose();
                        entry.Dispose();
                    }
                }
                catch
                {
                    dataGridView1.Rows.Add();
                    dataGridView1.Rows[index].Cells["Inv. č."].Value = "Nenalezeno žádné PC";
                }
            }
 

         }

        public Form1()
        {
            InitializeComponent();
        }

        public static bool PingHost(string nameOrAddress)
        {
            bool pingable = false;
            Ping pinger = new Ping();
            try
            {
                PingReply reply = pinger.Send(nameOrAddress);
                pingable = reply.Status == IPStatus.Success;
            }
            catch (PingException)
            {
                // Discard PingExceptions and return false;
            }
            return pingable;
        }



        public static string HostName2IP(string hostname)
        {
            // resolve the hostname into an iphost entry using the dns class
            try
            {
                XmlDocument doc = new XmlDocument();
                doc.Load("Config.xml");
                string DNSServer = doc.SelectSingleNode("/appSettings/configuration/DNSServer").InnerText;
                var client = new LookupClient(IPAddress.Parse(DNSServer));
                var result = client.Query(hostname, QueryType.A);
                string response="";
                foreach (var aRecord in result.Answers.ARecords())
                {
                    response = response + "," +aRecord.Address;
                }
                if (response == "") return "Neni v DNS";
                return response.Remove(0, 1);
            }
            catch
            {
                return "Neni v DNS";
            }
        }

        public static string EncryptString(string plainText, string sharedSecret)
        {
            string result = null;
            RijndaelManaged aesAlg = null;

            try
            {
                Rfc2898DeriveBytes key = new Rfc2898DeriveBytes(sharedSecret, _salt);
                aesAlg = new RijndaelManaged();
                aesAlg.Key = key.GetBytes(aesAlg.KeySize / 8);

                ICryptoTransform encryptor = aesAlg.CreateEncryptor(aesAlg.Key, aesAlg.IV);

                using (MemoryStream msEncrypt = new MemoryStream())
                {
                    msEncrypt.Write(BitConverter.GetBytes(aesAlg.IV.Length), 0, sizeof(int));
                    msEncrypt.Write(aesAlg.IV, 0, aesAlg.IV.Length);
                    using (CryptoStream csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write))
                    {
                        using (StreamWriter swEncrypt = new StreamWriter(csEncrypt))
                        {
                            swEncrypt.Write(plainText);
                        }
                    }
                    result = Convert.ToBase64String(msEncrypt.ToArray());
                }
            }
            finally
            {
                if (aesAlg != null)
                    aesAlg.Clear();
            }

            return result;
        }

        public static string DecryptString(string cipherText, string sharedSecret)
        {
            RijndaelManaged aesAlg = null;
            string result = null;

            try
            {
                Rfc2898DeriveBytes key = new Rfc2898DeriveBytes(sharedSecret, _salt);
                byte[] bytes = Convert.FromBase64String(cipherText);
                using (MemoryStream msDecrypt = new MemoryStream(bytes))
                {
                    aesAlg = new RijndaelManaged();
                    aesAlg.Key = key.GetBytes(aesAlg.KeySize / 8);
                    aesAlg.IV = ReadByteArray(msDecrypt);
                    ICryptoTransform decryptor = aesAlg.CreateDecryptor(aesAlg.Key, aesAlg.IV);
                    using (CryptoStream csDecrypt = new CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read))
                    {
                        using (StreamReader srDecrypt = new StreamReader(csDecrypt))
                        {
                            result = srDecrypt.ReadToEnd();
                        }
                    }
                }
            }
            finally
            {
                if (aesAlg != null)
                    aesAlg.Clear();
            }

            return result;
        }

        private static byte[] ReadByteArray(Stream s)
        {
            byte[] rawLength = new byte[sizeof(int)];
            if (s.Read(rawLength, 0, rawLength.Length) != rawLength.Length)
            {
                throw new SystemException("Stream did not contain properly formatted byte array");
            }

            byte[] buffer = new byte[BitConverter.ToInt32(rawLength, 0)];
            if (s.Read(buffer, 0, buffer.Length) != buffer.Length)
            {
                throw new SystemException("Did not read byte array properly");
            }

            return buffer;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (Control.ModifierKeys == Keys.Shift)
            {
                ReadPC();
            }
            else if (Control.ModifierKeys == Keys.Control)
            {
                ClearGrid();
            }
            else
            {
                RefreshOU();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load("Config.xml");
            string ADServer = doc.SelectSingleNode("/appSettings/configuration/ADServer").InnerText;
            string ADSearch = doc.SelectSingleNode("/appSettings/configuration/ADSearch").InnerText;
            string ADUser = doc.SelectSingleNode("/appSettings/configuration/ADUser").InnerText;
            //string ADPassword = doc.SelectSingleNode("/appSettings/configuration/ADPassword").InnerText;
            string ADPassword = DecryptString(doc.SelectSingleNode("/appSettings/configuration/ADPassword").InnerText, configPassword);
            string ADDomain = doc.SelectSingleNode("/appSettings/configuration/ADDomain").InnerText;
            bool ADSsso = Convert.ToBoolean(doc.SelectSingleNode("/appSettings/configuration/ADsso").InnerText);
            if (ADSsso)
            {
                DirectoryEntry entry = new DirectoryEntry("LDAP://" + ADServer + "/OU=" + comboBox1.Text + "," + ADSearch);
                DirectorySearcher mySearcher = new DirectorySearcher(entry);
                mySearcher.Filter = ("(objectCategory=Computer)");
                mySearcher.SizeLimit = int.MaxValue;
                mySearcher.PageSize = int.MaxValue;
                int index = 0;
                ClearGrid();
                dataGridView1.ColumnCount = 3;
                dataGridView1.ColumnHeadersVisible = true;



                // Set the column header style.
                DataGridViewCellStyle columnHeaderStyle = new DataGridViewCellStyle();

                columnHeaderStyle.BackColor = Color.Beige;
                columnHeaderStyle.Font = new Font("Verdana", 10, FontStyle.Bold);
                dataGridView1.ColumnHeadersDefaultCellStyle = columnHeaderStyle;

                // Set the column header names.
                dataGridView1.Columns[0].Name = "Inv. č.";
                dataGridView1.Columns[1].Name = "IP";
                dataGridView1.Columns[2].Name = "Ping";

                try
                {
                    SearchResultCollection results = mySearcher.FindAll();
                    {



                        foreach (SearchResult resEnt in mySearcher.FindAll())
                        {
                            string PCName = resEnt.GetDirectoryEntry().Name;


                            dataGridView1.Rows.Add();
                            dataGridView1.Rows[index].Cells["Inv. č."].Value = PCName.Remove(0, 3);
                            dataGridView1.Rows[index].Cells["IP"].Value = HostName2IP(PCName.Remove(0, 3) + "." + ADDomain);
                            if (HostName2IP(PCName.Remove(0, 3) + "." + ADDomain) != "Neni v DNS")
                            {
                                bool live = PingHost(PCName.Remove(0, 3) + "." + ADDomain);
                                if (live == true)
                                {
                                    dataGridView1.Rows[index].Cells["Ping"].Value = "OK";
                                    dataGridView1.Rows[index].DefaultCellStyle.BackColor = Color.Green;
                                }
                                else
                                {
                                    dataGridView1.Rows[index].Cells["Ping"].Value = "Vypnute";
                                    dataGridView1.Rows[index].DefaultCellStyle.BackColor = Color.Red;
                                }
                            }
                            else
                            {
                                dataGridView1.Rows[index].Cells["Ping"].Value = "Nelze pingnout";
                                dataGridView1.Rows[index].DefaultCellStyle.BackColor = Color.Yellow;
                            }
                            index++;
                        }

                        mySearcher.Dispose();
                        entry.Dispose();
                    }
                }
                catch
                {
                    dataGridView1.Rows.Add();
                    dataGridView1.Rows[index].Cells["Inv. č."].Value = "Nenalezeno žádné PC";
                }
            }
            else
            {
                DirectoryEntry entry = new DirectoryEntry("LDAP://" + ADServer + "/OU=" + comboBox1.Text + "," + ADSearch, ADUser, ADPassword);
                DirectorySearcher mySearcher = new DirectorySearcher(entry);
                mySearcher.Filter = ("(objectCategory=Computer)");
                mySearcher.SizeLimit = int.MaxValue;
                mySearcher.PageSize = int.MaxValue;
                int index = 0;
                ClearGrid();
                dataGridView1.ColumnCount = 3;
                dataGridView1.ColumnHeadersVisible = true;



                // Set the column header style.
                DataGridViewCellStyle columnHeaderStyle = new DataGridViewCellStyle();

                columnHeaderStyle.BackColor = Color.Beige;
                columnHeaderStyle.Font = new Font("Verdana", 10, FontStyle.Bold);
                dataGridView1.ColumnHeadersDefaultCellStyle = columnHeaderStyle;

                // Set the column header names.
                dataGridView1.Columns[0].Name = "Inv. č.";
                dataGridView1.Columns[1].Name = "IP";
                dataGridView1.Columns[2].Name = "Ping";

                try
                {
                    SearchResultCollection results = mySearcher.FindAll();
                    {



                        foreach (SearchResult resEnt in mySearcher.FindAll())
                        {
                            string PCName = resEnt.GetDirectoryEntry().Name;


                            dataGridView1.Rows.Add();
                            dataGridView1.Rows[index].Cells["Inv. č."].Value = PCName.Remove(0, 3);
                            dataGridView1.Rows[index].Cells["IP"].Value = HostName2IP(PCName.Remove(0, 3) + "." + ADDomain);
                            if (HostName2IP(PCName.Remove(0, 3) + "." + ADDomain) != "Neni v DNS")
                            {
                                bool live = PingHost(PCName.Remove(0, 3) + "." + ADDomain);
                                if (live == true)
                                {
                                    dataGridView1.Rows[index].Cells["Ping"].Value = "OK";
                                    dataGridView1.Rows[index].DefaultCellStyle.BackColor = Color.Green;
                                }
                                else
                                {
                                    dataGridView1.Rows[index].Cells["Ping"].Value = "Vypnute";
                                    dataGridView1.Rows[index].DefaultCellStyle.BackColor = Color.Red;
                                }
                            }
                            else
                            {
                                dataGridView1.Rows[index].Cells["Ping"].Value = "Nelze pingnout";
                                dataGridView1.Rows[index].DefaultCellStyle.BackColor = Color.Yellow;
                            }
                            index++;
                        }

                        mySearcher.Dispose();
                        entry.Dispose();
                    }
                }
                catch
                {
                    dataGridView1.Rows.Add();
                    dataGridView1.Rows[index].Cells["Inv. č."].Value = "Nenalezeno žádné PC";
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            
    
          
        }

        private void button4_Click(object sender, EventArgs e)
        {
            // creating Excel Application
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            // creating new WorkBook within Excel application
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            // creating new Excelsheet in workbook
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            // see the excel sheet behind the program
            app.Visible = true;
            // get the reference of first sheet. By default its name is Sheet1.
            // store its reference to worksheet
            worksheet = workbook.Sheets["List1"];
            worksheet = workbook.ActiveSheet;
            // changing the name of active sheet
            worksheet.Name = "Ping";
            // storing header part in Excel
            for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
            {
                worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
            }
            // storing Each row and column value to excel sheet
            for (int i = 0; i < dataGridView1.Rows.Count-1; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                }
            }
            // save the application
            //workbook.SaveAs("c:\\output.xls", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            // Exit from the application
            //app.Quit();
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            string filePath = openFileDialog1.FileName;
            string extension = Path.GetExtension(filePath);
            //string header = rbHeaderYes.Checked ? "YES" : "NO";
            string header = "YES";
            string conStr;
            //string sheetName;

            conStr = string.Empty;
            switch (extension)
            {

                case ".xls": //Excel 97-03
                    conStr = string.Format(Excel03ConString, filePath, header);
                    break;

                case ".xlsx": //Excel 07
                    conStr = string.Format(Excel07ConString, filePath, header);
                    break;
            }

            ClearGrid();

            //Get the name of the First Sheet.
            /*using (OleDbConnection con = new OleDbConnection(conStr))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    cmd.Connection = con;
                    con.Open();
                    DataTable dtExcelSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                    MessageBox.Show(sheetName);
                    con.Close();
                }
            }*/

            //Read Data from the First Sheet.
            using (OleDbConnection con = new OleDbConnection(conStr))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    using (OleDbDataAdapter oda = new OleDbDataAdapter())
                    {
                        DataTable dt = new DataTable();
                        cmd.CommandText = "SELECT * From [Ping$]";
                        cmd.Connection = con;
                        con.Open();
                        oda.SelectCommand = cmd;
                        oda.Fill(dt);
                        con.Close();

                        //Populate DataGridView.
                        dataGridView1.DataSource = dt;
                    }
                }
            }
            // Set the column header style.
            DataGridViewCellStyle columnHeaderStyle = new DataGridViewCellStyle();

            columnHeaderStyle.BackColor = Color.Beige;
            columnHeaderStyle.Font = new Font("Verdana", 10, FontStyle.Bold);
            dataGridView1.ColumnHeadersDefaultCellStyle = columnHeaderStyle;
            // Set the column header names.
            dataGridView1.Columns[0].Name = "Inv. č.";
            dataGridView1.Columns[1].Name = "IP";
            dataGridView1.Columns[2].Name = "Ping";
            dataGridView1.Columns[0].HeaderText = "Inv. č.";
            dataGridView1.Columns[1].HeaderText = "IP";
            dataGridView1.Columns[2].HeaderText = "Ping";
            
            int rowCount2 = dataGridView1.Rows.Count;
            for (int n = 0; n < rowCount2-1; n++)
            {


                dataGridView1.Rows[n].Cells["IP"].Value = HostName2IP(dataGridView1.Rows[n].Cells["Inv. č."].Value + ".fnol.loc");
                if (HostName2IP(dataGridView1.Rows[n].Cells["IP"].Value.ToString()) != "Neni v DNS")
                {
                    bool live = PingHost(dataGridView1.Rows[n].Cells["IP"].Value.ToString());
                    if (live == true)
                    {
                        dataGridView1.Rows[n].Cells["Ping"].Value = "OK";
                        dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Green;
                    }
                    else
                    {
                        dataGridView1.Rows[n].Cells["Ping"].Value = "Vypnute";
                        dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Red;
                    }
                }
                else
                {
                    dataGridView1.Rows[n].Cells["Ping"].Value = "Nelze pingnout";
                    dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Yellow;
                }
                
            }
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            RefreshOU();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ReadPC();
        }

      
        
    }
}
