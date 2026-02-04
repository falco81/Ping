using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace Pinger
{
    static class Program
    {
        // use console from another process
        //[System.Runtime.InteropServices.DllImport("kernel32.dll")]
        //static extern bool AttachConsole(int procId);

        //private const int ATTACH_PARENT_PROCESS = -1;

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {

            if (Array.Exists(args, s => s.Equals("-c")))
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Config());

/*              // redirect console output to parent process;
                // must be before any calls to Console.WriteLine()
                AttachConsole(ATTACH_PARENT_PROCESS);

                // process arguments and use console (line below just for sample debug)
                Console.WriteLine(string.Join(",", args));
                string crypttext = Form1.EncryptString(args[1], Form1.configPassword);
                Console.WriteLine("Password saved.");
             
                XmlDocument doc = new XmlDocument();
                doc.Load("Config.xml");
                XmlNode node = doc.SelectSingleNode("/appSettings/configuration/ADPassword");
                node.InnerText = crypttext;
                doc.Save("Config.xml");
                //Console.ReadKey();
*/
            }
            else
            {
              //  if (args.Length > 0)
              //  {
                    //MessageBox.Show("Arguments: " + string.Join(",", args));
              //  }

                // show GUI like
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Form1());

            }
        }
    }
}
