using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelMacroRun
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();

            txtTartDir.Text = "";
        }

        private void btnSelectDir_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txtTartDir.Text = folderBrowserDialog1.SelectedPath;
            } else {
                return;
            }
            // Set Items on ListBox
            string[] files = System.IO.Directory.GetFiles(@txtTartDir.Text, "*.xlsm", System.IO.SearchOption.AllDirectories);
            lboxFileList.Items.AddRange(files);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Run OK?",
                                                  "Question",
                                                   MessageBoxButtons.YesNoCancel,
                                                   MessageBoxIcon.Exclamation,
                                                   MessageBoxDefaultButton.Button2);

            if (result != DialogResult.Yes)
            {
                return;
            }

            foreach (var listitem in lboxFileList.Items)
            {
                System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo();
                psi.FileName = "excel.exe";
                psi.Arguments = listitem.ToString();
                System.Diagnostics.Process p = System.Diagnostics.Process.Start(psi);

                Console.WriteLine(DateTime.Now + ":" + listitem.ToString());
                
                //wait time
                System.Threading.Thread.Sleep(3000); //1sec = 1000

                //check end file
                while (!System.IO.File.Exists(listitem.ToString().Replace(".xlsm",".ok.log")))
                {
                    //wait time
                    System.Threading.Thread.Sleep(3000); //1sec = 1000
                }
            }

            MessageBox.Show("Completed!");
        }
    }
}
