using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PSMOpenFolders
{
    public partial class Form1 : System.Windows.Forms.Form
    {
        public Form1()
        {
            InitializeComponent();
            this.ActiveControl = txtJobNo; // <--- set the active control
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //To make combobox non editable
            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
            //Set preferred index to show as default value
            comboBox1.SelectedIndex = 0;
        }

        private void btnOpenFolder_Click(object sender, EventArgs e)
        {
            OpenFolder();
        }

        public string folderPath(int jobNum)
        {
            string PSM = "PSM";
            string jobNumStr;
            if (jobNum < 10)
            {
                PSM = PSM + "00";
            }
            else if (jobNum < 100)
            {
                PSM = PSM + "0";
            }

            jobNumStr = PSM + jobNum.ToString();

            int driveIndex = comboBox1.SelectedIndex;

            string naspath;
            if (driveIndex == 0)
            {
                naspath = @""; 
            }
            else
            {
                naspath = @"\\files-server\LocalSoftwareDrive";
            }
            
            string fpath = null;
            // Folder number divisions
            int[] NumIncr = { 0, 1000, 1400, 1800, 2200, 2600, 3000, 4000, 5000};
            string[] NumIncrStr = { @"U:\", @"T:\", @"S:\", @"R:\", @"Q:\", @"P:\", @"O:\", @"N:\"};
            string thousands = null;
            for (int i = 0; i < NumIncr.Length-1; i+=1)
            {
                if (jobNum >= NumIncr[i] && jobNum < NumIncr[i + 1])
                {
                    thousands = NumIncrStr[i];
                }
            }
            string fpath1 = Path.Combine(naspath, thousands);
            string fpath2 = Path.Combine(fpath1, jobNumStr);
            if (Directory.Exists(fpath2))
            {
                fpath = fpath2;
            }
            else
            {
                string fpath3 = Path.Combine(naspath, "3000 Expanded");
                string fpath4 = Path.Combine(fpath3, jobNumStr);
                if (Directory.Exists(fpath4))
                {
                    fpath = fpath4;
                }

            }

            return fpath;
        }

        public void OpenFolder()
        {
            int jobNum = 0;
            try
            {
                jobNum = Int32.Parse(txtJobNo.Text);
                if (jobNum >= 0 && jobNum < 5000)
                {
                    string fpath = folderPath(jobNum);
                    if (fpath == null)
                    {
                        MessageBox.Show("Folder does not exist");
                    }
                    else
                    {
                        System.Diagnostics.Process.Start(fpath);
                    }
                }
                else
                {
                    MessageBox.Show("Job number exceeds range");
                }
            }
            catch
            {
                MessageBox.Show("Please enter a job number");
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        public void OpenLink(string linkpath)
        {
            try
            {
                System.Diagnostics.Process.Start(linkpath);
            }
            catch (Exception err)
            {
                MessageBox.Show("Folder does not exist");
            }
        }
        private void btnStandards_Click(object sender, EventArgs e)
        {
            OpenLink(@"W:\Technical\Specifications Standards and Guidelines\");
        }

        private void btnLibrary_Click(object sender, EventArgs e)
        {
            OpenLink(@"W:\Technical\Library\");
        }

        private void btnSoftwareeng_Click(object sender, EventArgs e)
        {
            OpenLink(@"W:\Technical\Software_Engineering\");
        }

        private void btnSoftwarewindows_Click(object sender, EventArgs e)
        {
            OpenLink(@"W:\Technical\Software_Windows\");
        }

        private void btnTemplates_Click(object sender, EventArgs e)
        {
            OpenLink(@"W:\Public\Templates\");
        }

        private void btnPsmtools_Click(object sender, EventArgs e)
        {
            OpenLink(@"W:\Technical\PSM Tools\");
        }

        private void txtJobNo_TextChanged(object sender, EventArgs e)
        {
            bool enteredLetter = false;
            Queue<char> text = new Queue<char>();
            foreach (var ch in this.txtJobNo.Text)
            {
                if (char.IsDigit(ch))
                {
                    text.Enqueue(ch);
                }
                else
                {
                    enteredLetter = true;
                }
            }

            if (enteredLetter)
            {
                StringBuilder sb = new StringBuilder();
                while (text.Count > 0)
                {
                    sb.Append(text.Dequeue());
                }

                this.txtJobNo.Text = sb.ToString();
                this.txtJobNo.SelectionStart = this.txtJobNo.Text.Length;
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        
        public void OpenWeb(string link)
        {
            try
            {
                System.Diagnostics.Process.Start(link);
            }
            catch (Exception err)
            {
                MessageBox.Show("Cannot open webpage");
            }
        }

        private void btnLicenses_Click(object sender, EventArgs e)
        {
            OpenWeb("http://ls.psm.com.au:88/status/");
        }

        private void btnSharepoint_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("iexplore", "https://psm.deltekfirst.com/PSMClient/DeltekVision.application");
        }

        private void btnHelp_Click(object sender, EventArgs e)
        {
            OpenWeb("https://psmau.sharepoint.com/sites/PSMHelpPages/");
        }

        private void btnBrisbane_Click(object sender, EventArgs e)
        {
            
            if (Directory.Exists(@"\\10.37.80.51\Bris_Files"))
            {
                OpenLink(@"\\10.37.80.51\Bris_Files");
            }
            else
            {
                // Remove network drive mapping
                System.Diagnostics.Process.Start("net.exe", @"use B: /DELETE /YES").WaitForExit();
                // Create network drive mapping
                System.Diagnostics.Process.Start("net.exe", @"USE B: ""\\10.37.80.51\Bris_files /y "" /persistent:yes /user:Brisbaneuser PSM1993").WaitForExit();
                OpenLink(@"\\10.37.80.51\Bris_Files");
            }

        }

        private void btnMining_Click(object sender, EventArgs e)
        {
            OpenLink(@"\\Files-server\Mining\");
        }

        private void btnVulcan_Click(object sender, EventArgs e)
        {
            OpenLink(@"Z:\");
        }

        private void btnLocal_Click(object sender, EventArgs e)
        {
            OpenLink(@"Y:\");
        }

        private void btnTechnical_Click(object sender, EventArgs e)
        {
            OpenLink(@"W:\Technical\");
        }

        private void btnPublic_Click(object sender, EventArgs e)
        {
            OpenLink(@"W:\Public\");
        }
    }
}
