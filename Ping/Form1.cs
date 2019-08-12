using System;
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
            string ADPassword = CryptUtils.DecryptString(doc.SelectSingleNode("/appSettings/configuration/ADPassword").InnerText, CryptUtils.configPassword);
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
            string ADPassword = CryptUtils.DecryptString(doc.SelectSingleNode("/appSettings/configuration/ADPassword").InnerText, CryptUtils.configPassword);
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
            string ADPassword = CryptUtils.DecryptString(doc.SelectSingleNode("/appSettings/configuration/ADPassword").InnerText, CryptUtils.configPassword);
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
                        progressBar1.Maximum = mySearcher.FindAll().Count;

                        foreach (SearchResult resEnt in mySearcher.FindAll())
                        {
                            progressBar1.Value++;
                            string PCName = resEnt.GetDirectoryEntry().Name;


                            dataGridView1.Rows.Add();
                            dataGridView1.Rows[index].Cells["Inv. č."].Value = PCName.Remove(0, 3);
                            dataGridView1.Rows[index].Cells["IP"].Value = HostName2IP(PCName.Remove(0, 3) + "." + ADDomain);
                            if (dataGridView1.Rows[index].Cells["IP"].Value.ToString() != "Neni v DNS")
                            {
                                bool live = PingHost(dataGridView1.Rows[index].Cells["IP"].Value.ToString());
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
                        progressBar1.Value = 0;
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

                        progressBar1.Maximum = mySearcher.FindAll().Count;

                        foreach (SearchResult resEnt in mySearcher.FindAll())
                        {
                            progressBar1.Value++;
                            string PCName = resEnt.GetDirectoryEntry().Name;


                            dataGridView1.Rows.Add();
                            dataGridView1.Rows[index].Cells["Inv. č."].Value = PCName.Remove(0, 3);
                            dataGridView1.Rows[index].Cells["IP"].Value = HostName2IP(PCName.Remove(0, 3) + "." + ADDomain);
                            if (dataGridView1.Rows[index].Cells["IP"].Value.ToString() != "Neni v DNS")
                            {
                                bool live = PingHost(dataGridView1.Rows[index].Cells["IP"].Value.ToString());
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
                        progressBar1.Value = 0;
                    }
                }
                catch
                {
                    dataGridView1.Rows.Add();
                    dataGridView1.Rows[index].Cells["Inv. č."].Value = "Nenalezeno žádné PC";
                }
            }
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
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                XmlDocument doc = new XmlDocument();
                doc.Load("Config.xml");
                string ADDomain = doc.SelectSingleNode("/appSettings/configuration/ADDomain").InnerText;
                string filePath = openFileDialog1.FileName;
                string extension = Path.GetExtension(filePath);
                ClearGrid();

                // Get the Excel application object.
                Excel.Application excel_app = new Excel.Application();

                // Make Excel visible (optional).
                excel_app.Visible = true;

                // Open the workbook read-only.
                Excel.Workbook workbook = excel_app.Workbooks.Open(
                    filePath,
                    Type.Missing, true, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);

                // Get the first worksheet.
                Excel.Worksheet sheet = (Excel.Worksheet)workbook.Sheets[1];

                // Get the used range.
                Excel.Range used_range = sheet.UsedRange;

                // Get the maximum row and column number.
                int max_row = used_range.Rows.Count;
                int max_col = used_range.Columns.Count;

                // Get the sheet's values.
                object[,] values = (object[,])used_range.Value2;

                // Get the column titles.
                SetGridColumns(dataGridView1, values, max_col);

                // Get the data.
                SetGridContents(dataGridView1, values, max_row, max_col);

                // Close the workbook without saving changes.
                workbook.Close(false, Type.Missing, Type.Missing);

                // Close the Excel server.
                excel_app.Quit();



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
                progressBar1.Maximum = rowCount2;
                for (int n = 0; n < rowCount2 - 1; n++)
                {
                    progressBar1.Value++;

                    dataGridView1.Rows[n].Cells["IP"].Value = HostName2IP(dataGridView1.Rows[n].Cells["Inv. č."].Value + "." + ADDomain);
                    if (dataGridView1.Rows[n].Cells["IP"].Value.ToString() != "Neni v DNS")
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
                progressBar1.Value = 0;

            }
        }

        private void SetGridColumns(DataGridView dgv,
           object[,] values, int max_col)
        {
            dataGridView1.Columns.Clear();

            // Get the title values.
            for (int col = 1; col <= max_col; col++)
            {
                string title = (string)values[1, col];
                dgv.Columns.Add("col_" + title, title);
            }
        }

        // Set the grid's contents.
        private void SetGridContents(DataGridView dgv,
            object[,] values, int max_row, int max_col)
        {
            // Copy the values into the grid.
            for (int row = 2; row <= max_row; row++)
            {
                object[] row_values = new object[max_col];
                for (int col = 1; col <= max_col; col++)
                    row_values[col - 1] = values[row, col];
                dgv.Rows.Add(row_values);
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

        private void Button5_Click(object sender, EventArgs e)
        {
            int rowCount2 = dataGridView1.Rows.Count;
            progressBar1.Maximum = rowCount2;
            for (int n = 0; n < rowCount2 - 1; n++)
            {
                progressBar1.Value++;
                XmlDocument doc = new XmlDocument();
                doc.Load("Config.xml");
                string ADDomain = doc.SelectSingleNode("/appSettings/configuration/ADDomain").InnerText;

                dataGridView1.Rows[n].Cells["IP"].Value = HostName2IP(dataGridView1.Rows[n].Cells["Inv. č."].Value + "." + ADDomain);
                if (dataGridView1.Rows[n].Cells["IP"].Value.ToString() != "Neni v DNS")
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
            progressBar1.Value=0;
        }

        private void Button6_Click(object sender, EventArgs e)
        {
            ClearGrid();
        }
    }
}
