using System;
using System.Collections.Generic;
using System.Configuration;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Management;
using System.Net.NetworkInformation;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace COPS
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            this.Text += Application.ProductVersion.ToString();
            tb_RemoteDIR.Text = "C:\\442x442";
            tb_LogLocation.Text = Environment.CurrentDirectory + "\\Logging";
            tb_LocalDIR.Text = Environment.CurrentDirectory + "\\DeploymentPrep";
            if (!File.Exists(tb_LocalDIR.Text))
            {
                Directory.CreateDirectory(tb_LocalDIR.Text);
            }
            if (!Directory.Exists(Environment.CurrentDirectory + "\\Utilities"))
            {
                Directory.CreateDirectory(Environment.CurrentDirectory + "\\Utilities");
            }
            if (!Directory.Exists(Environment.CurrentDirectory + "\\Configs"))
            {
                Directory.CreateDirectory(Environment.CurrentDirectory + "\\Configs");
            }
            if (!Directory.Exists(Environment.CurrentDirectory + "\\Logging"))
            {
                Directory.CreateDirectory(Environment.CurrentDirectory + "\\Logging");
            }
            if (!Directory.Exists(Environment.CurrentDirectory + "\\Documentation"))
            {
                Directory.CreateDirectory(Environment.CurrentDirectory + "\\Documentation");
            }
            if (!File.Exists(Environment.CurrentDirectory + "\\Configs\\CONFIG-InstallerPresets.txt"))
            {
                File.Create(Environment.CurrentDirectory + "\\Configs\\CONFIG-InstallerPresets.txt");
            }
            if (!File.Exists(Environment.CurrentDirectory + "\\Configs\\CONFIG-CyberOrderTypes.txt"))
            {
                File.Create(Environment.CurrentDirectory + "\\Configs\\CONFIG-CyberOrderTypes.txt");
            }
            if (!File.Exists(Environment.CurrentDirectory + "\\Utilities\\Psexec64.exe"))
            {
                MessageBox.Show("PSExec is missing from: " + Environment.NewLine + Environment.CurrentDirectory + @"\Utilities",
                                "Required File Missing!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            Thread.Sleep(500);
            foreach (string item in File.ReadAllLines(Environment.CurrentDirectory + "\\Configs\\CONFIG-InstallerPresets.txt"))
            {
                cb_Presets.Items.Add(item);
            }
            foreach (string item in File.ReadAllLines(Environment.CurrentDirectory + "\\Configs\\CONFIG-CyberOrderTypes.txt"))
            {
                cb_OrderType.Items.Add(item);
            }
            foreach(string filename in Directory.GetFiles(Environment.CurrentDirectory + "\\Documentation"))
            {
                lb_Documentation.Items.Add(filename.ToString().Split(@"\").Last());
            }
            foreach (string filename in Directory.GetFiles(Environment.CurrentDirectory + "\\Configs"))
            {
                if (filename.Contains("CONFIG-"))
                {
                    lb_Documentation.Items.Add(filename.ToString().Split(@"\").Last());
                }
            }

        }
        private string CalcauteInstallerBytes(double filelength)
        {
            int counter = 0;
            while (true)
            {
                if (Math.Round(filelength,0).ToString().Length >= 4)
                {
                    filelength /= 1024;
                    counter++;
                }
                else
                {
                    break;
                }
            }
            string suffix = "";
            switch (counter)
            {
                case 0:
                    suffix = " Bs";
                    break;
                case 1:
                    suffix = " KBs";
                    break;
                case 2:
                    suffix = " MBs";
                    break;
                case 3:
                    suffix = " GBs";
                    break;
                case 4:
                    suffix = " TBs";
                    break;
                default:
                    suffix = " File Too Large!";
                    break;
            }
            return (Math.Round(filelength,2)).ToString() + suffix;
        }
        private void tb_Installer_TextChanged(object sender, EventArgs e)
        {
            if(File.Exists(Global.strFullPathToInstaller))
            {
                FileInfo fi = new FileInfo(Global.strFullPathToInstaller);
                long filelength = fi.Length;
                label2.Text = CalcauteInstallerBytes(filelength);
            }
            else
            {
                label2.Text = "No File Selected";
                return;
            }
        }
        private void bnt_SelectInstaller_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.InitialDirectory = tb_LocalDIR.Text;
            if(ofd.ShowDialog() == DialogResult.OK)
            {
                if (File.Exists(ofd.FileName))
                {
                    Global.strFullPathToInstaller = ofd.FileName;
                    tb_Installer.Text = ofd.FileName.Split("\\").Last();
                    FileInfo fi = new FileInfo(ofd.FileName);
                    long filelength = fi.Length;
                    label2.Text = CalcauteInstallerBytes(filelength);
                }
            }
            ofd.Dispose();
        }
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            string f = tb_Installer.Text;
            if (f == null || f.Length == 0 || f == "")
            {
                return;
            }
            tb_Parameters.Text = cb_Presets.SelectedItem.ToString().Split("|").Last().Trim().Replace("$PATH", tb_RemoteDIR.Text).Replace("$FILENAME", f.Split("\\").Last());
            if (cb_Presets.Text.ToString().Contains("CEDS"))
            {
                DirectoryInfo di = new DirectoryInfo(Global.strFullPathToInstaller.Substring(0, Global.strFullPathToInstaller.LastIndexOf(@"\")));
                string strParentFolder = di.Name;
                FileInfo[] fi = di.GetFiles("*.*", SearchOption.AllDirectories);
                double foldersize = 0;
                for (int i = 0; i < fi.Count(); i++)
                {
                    foldersize += fi[i].Length;
                }
                label2.Text = CalcauteInstallerBytes(foldersize);
                tb_Parameters.Text = tb_Parameters.Text.Replace("Install.cmd", strParentFolder + @"\Install.cmd");
                Application.DoEvents();
            }
        }
        private void aboutToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Version: " + Application.ProductVersion.ToString() + Environment.NewLine +
                            "Date: 1 Jan 2022" + Environment.NewLine +
                            "Created with Visual Studio | DOT NET 6", "Product of the 442 Hotline",MessageBoxButtons.OK,MessageBoxIcon.Information);
        }
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        private void btn_LocalDIRDialog_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if(fbd.ShowDialog() == DialogResult.OK)
            {
                tb_LocalDIR.Text = fbd.SelectedPath;
            }
            else
            {
                tb_LocalDIR.Text = "";
            }
            Application.DoEvents();
        }
        private void btn_RemoteDIRDialog_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                tb_RemoteDIR.Text = fbd.SelectedPath;
            }
            else
            {
                tb_RemoteDIR.Text = "";
            }
            Application.DoEvents();
        }
        private void btn_GetComputerList_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.InitialDirectory = Environment.CurrentDirectory;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                if (File.Exists(ofd.FileName))
                {
                    tb_CompList.Text = ofd.FileName;
                    list_Computers.Lines = File.ReadAllLines(ofd.FileName);
                    
                    Application.DoEvents();
                }
            }
        }
        private void reloadConfigsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if(File.Exists(Environment.CurrentDirectory + "\\Configs\\CONFIG-InstallerPresets.txt"))
            {
                cb_Presets.Items.Clear();
                foreach (string item in File.ReadAllLines(Environment.CurrentDirectory + "\\Configs\\CONFIG-InstallerPresets.txt"))
                {
                    cb_Presets.Items.Add(item);
                }
            }
            if (File.Exists(Environment.CurrentDirectory + "\\Configs\\CONFIG-CyberOrderTypes.txt"))
            {
                cb_OrderType.Items.Clear();
                foreach (string item in File.ReadAllLines(Environment.CurrentDirectory + "\\Configs\\CONFIG-CyberOrderTypes.txt"))
                {
                    cb_OrderType.Items.Add(item);
                }
            }
            lb_Documentation.Items.Clear();
            foreach (string filename in Directory.GetFiles(Environment.CurrentDirectory + "\\Documentation"))
            {
                lb_Documentation.Items.Add(filename.ToString().Split(@"\").Last());
            }
            foreach (string filename in Directory.GetFiles(Environment.CurrentDirectory + "\\Configs"))
            {
                lb_Documentation.Items.Add(filename.ToString().Split(@"\").Last());
            }

            if (File.Exists(tb_CompList.Text))
            {
                list_Computers.Lines = File.ReadAllLines(tb_CompList.Text);
            }           
        }

        private void btn_DeployPreCheck_Click(object sender, EventArgs e)
        {
            PreDeploymentChecks PDC = new PreDeploymentChecks();
            PDC = ControlsCheck(PDC);
            foreach (string strline in File.ReadAllLines(Environment.CurrentDirectory + "\\Configs\\CONFIG-CyberOrderTypes.txt"))
            {
                if (cb_OrderType.Text.ToString().Trim().Contains(strline.Trim()))
                {
                    PDC.CyberOrderCheck = true;
                    break;
                }
            }
            if (File.Exists(Global.strFullPathToInstaller))
            {
                PDC.InstallerCheck = true;
            }
            if(list_Computers.Text.Length != 0)
            {
                PDC.ComputerListCheck = true;
            }
            if (PDC.RunChecklist())
            {
                btn_Deploy.Enabled = true;
                btn_Deploy.BackColor = Color.Green;
            }
            else
            {
                btn_Deploy.Enabled = false;
                btn_Deploy.BackColor = Color.Yellow;
            }
        }
        /// <summary>
        /// Checks Required Textboxes and Comboxes for length
        /// </summary>
        /// <param name="PDC"></param>
        private PreDeploymentChecks ControlsCheck(PreDeploymentChecks PDC)
        {
            if (Utility.ControlValid(tb_OrderID))
            {
                PDC.CyberOrderIDCheck = true;
            }
            if (Utility.ControlValid(cb_Presets))
            {
                PDC.PresetsCheck = true;
            }
            if (Utility.ControlValid(tb_LocalDIR))
            {
                PDC.LocalDIRCheck = true;
            }
            if (Utility.ControlValid(tb_RemoteDIR))
            {
                PDC.RemoteDIRCheck = true;
            }
            if (Utility.ControlValid(tb_Parameters))
            {
                PDC.ParametersCheck = true;
            }
            return PDC;
        }
        private void btn_Deploy_Click(object sender, EventArgs e)
        {
            tabControl1.SelectTab("tabPage2");
            Application.DoEvents();
            DataTable dt = new DataTable();
            dt.Columns.Add("ComputerName", Type.GetType("System.String"));
            dt.Columns.Add("IPv4", Type.GetType("System.String"));
            dt.Columns.Add("PingStatus", Type.GetType("System.String"));
            dt.Columns.Add("AdminShareResult", Type.GetType("System.String"));
            dt.Columns.Add("PSExecErrorCode", Type.GetType("System.String"));
            List<Task>TaskList = new List<Task>();
            TaskFactory factory = new TaskFactory();
            int intCounter = 0;
            int intRunningTaskCounter = 0;
            toolStripProgressBar1.Maximum = list_Computers.Lines.Count();
            foreach (string strComputer in list_Computers.Lines)
            {
                intCounter++;
                intRunningTaskCounter++;
                toolStripProgressBar1.Value = intCounter;
                toolStripStatusLabel3.Text = intRunningTaskCounter.ToString() + "Tasks Running";
                toolStripStatusLabel1.Text = intCounter.ToString() + " of " + list_Computers.Lines.Count().ToString() + " task[s] queued...";
                Application.DoEvents();

                while (intRunningTaskCounter > num_MaxTasks.Value)
                {
                    for (int i = 0; i < TaskList.Count; i++)
                    {
                        Task task = TaskList[i];
                        if (task.IsCompleted)
                        {
                            intRunningTaskCounter--;
                        }
                    }
                    Application.DoEvents();
                }
                string A = new string(strComputer);
                string B = new string(Global.strFullPathToInstaller);
                string C = new string(cb_Presets.Text.ToString());
                string D = new string(tb_LocalDIR.Text.ToString());
                string E = new string(tb_RemoteDIR.Text.ToString());
                string F = new string(tb_Parameters.Text.ToString());
                Task<DeploymentResults> t = Task<DeploymentResults>.Factory.StartNew(() => TaskRoutine(A,B,C,D,E,F));
                TaskList.Add(t);
                Application.DoEvents();
            }

            while (TaskList.Count() > 0)
            {
                foreach (Task tt in TaskList)
                {
                    if (tt.IsCompleted)
                    {
                        DeploymentResults ddrr = ((Task<DeploymentResults>)tt).Result;
                        dt.Rows.Add(ddrr.strComputer, ddrr.strIPv4Address, ddrr.strPingStatus, ddrr.strAdminShareStatus, ddrr.strInstallResult);
                        TaskList.Remove(tt);
                        dataGridView1.DataSource = dt;
                        Application.DoEvents();
                        break;
                    }
                }
                toolStripStatusLabel3.Text = TaskList.Count.ToString() + " tasks left...";
                Application.DoEvents();
            }
            dataGridView1.DataSource = dt;
            toolStripProgressBar1.Value = 0;
            toolStripStatusLabel3.Text = "Ready";
            toolStripStatusLabel1.Text = "";
        }
        private DeploymentResults TaskRoutine(string strComputer, string strInstallFile, string strPresets, string strLocalDIR, string strRemoteDIR, string strParameters)
        {
            DeploymentResults DR = new DeploymentResults();
            PingOptions po = new PingOptions();
            Ping p = new Ping();
            po.Ttl = 1000;
            DR.strComputer = strComputer;
            PingReply pr = p.Send(strComputer);
            if (pr.Status == IPStatus.Success)
            {
                DR.strPingStatus = pr.Status.ToString();
                DR.strIPv4Address = pr.Address.ToString();
            }
            else
            {
                if(pr.Address.ToString() == "" || pr.Address.ToString() == null)
                {
                    DR.strIPv4Address = "CANNOT RESOLVE";
                }
                else
                {
                    DR.strIPv4Address = pr.Address.ToString();
                }
                DR.strPingStatus = pr.Status.ToString();
                DR.strAdminShareStatus = "FAILED";
                DR.strInstallResult = "FAILED";
                return DR;
            }
            if(Directory.Exists(@"\\" + strComputer + @"\c$\") && Directory.Exists(@"\\" + strComputer + @"\admin$\"))
            {
                DR.strAdminShareStatus = "Success";
            }
            else
            {
                DR.strAdminShareStatus = "FAILED";
                DR.strInstallResult = "FAILED";
                return DR;
            }
            Process proc = new Process();
            if (strPresets.Contains("CEDS"))
            {
                FileInfo fi = new FileInfo(strInstallFile);
                DIO dIO = new DIO();
                dIO.CopyFilesRecursively(fi.Directory.Parent.FullName.ToString(), strRemoteDIR);
                double LfolderSize = dIO.CalculateFolderSize(fi.Directory.Parent.FullName.ToString());
                while(LfolderSize != dIO.CalculateFolderSize(strRemoteDIR))
                {
                    Thread.Sleep(1000);
                }
            }
            else
            {
                FileInfo fi = new FileInfo(strInstallFile);
                File.Copy(strInstallFile, strRemoteDIR.Replace("C:", @"\\" + strComputer + @"\c$") + @"\" + strInstallFile.Split(@"\").Last(), true);
                while(true)
                {
                    if(File.Exists(strRemoteDIR.Replace("C:", @"\\" + strComputer + @"\c$") + @"\" + strInstallFile.Split(@"\").Last()))
                    {
                        FileInfo fo = new FileInfo(strRemoteDIR.Replace("C:", @"\\" + strComputer + @"\c$") + @"\" + strInstallFile.Split(@"\").Last());
                        if (fi.Length == fo.Length)
                        {
                            break;
                        }
                    }
                    Thread.Sleep(500);
                }
            }
            ProcessStartInfo psi = new ProcessStartInfo();
            psi.UseShellExecute = true;
            psi.FileName = Environment.CurrentDirectory + @"\Utilities\Psexec64.exe";
            psi.Arguments = @"\\" + strComputer + " -acceupteula -s " + strParameters;
            psi.WindowStyle = ProcessWindowStyle.Hidden;
            proc.StartInfo = psi;
            proc.Start();
            proc.WaitForExit();
            DR.strInstallResult = proc.ExitCode.ToString();

            if (strPresets.Contains("CEDS"))
            {
                FileInfo fi = new FileInfo(strParameters.Replace("cmd.exe /C ","").Trim());
                Directory.Delete(fi.Directory.FullName.ToString(), true);
            }
            else
            {
                File.Delete(strRemoteDIR.Replace("C:", @"\\" + strComputer + @"\c$") + @"\" + strInstallFile.Split(@"\").Last());
            }

            return DR;
        }
        private void btn_ExportClear_Click(object sender, EventArgs e)
        {
            string strFile = tb_LogLocation.Text + @"\" + Environment.UserName + "-" + cb_OrderType.Text.ToString() + "-" + tb_OrderID.Text.ToString() + "_" + DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss").ToString() + ".csv";
            StreamWriter file = new StreamWriter(strFile);
            file.WriteLine(dataGridView1.Columns[0].HeaderText.ToString() + "," +
                           dataGridView1.Columns[1].HeaderText.ToString() + "," +
                           dataGridView1.Columns[2].HeaderText.ToString() + "," +
                           dataGridView1.Columns[3].HeaderText.ToString() + "," +
                           dataGridView1.Columns[4].HeaderText.ToString());
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                file.WriteLine(dataGridView1.Rows[i].Cells[0].Value.ToString() + "," +
                               dataGridView1.Rows[i].Cells[1].Value.ToString() + "," +
                               dataGridView1.Rows[i].Cells[2].Value.ToString() + "," +
                               dataGridView1.Rows[i].Cells[3].Value.ToString() + "," +
                               dataGridView1.Rows[i].Cells[4].Value.ToString());
            }
            file.Close();
            MessageBox.Show("Log file has been exported to: " + Environment.NewLine +
                            tb_LogLocation.Text.ToString() + @"\" + cb_OrderType.Text.ToString() + "_" + DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss").ToString() + ".csv",
                            "Export Results", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void btn_LogLocation_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                tb_LogLocation.Text = fbd.SelectedPath;
            }
            else
            {
                tb_LogLocation.Text = "";
            }
            if (!Directory.Exists(tb_LogLocation.Text))
            {
                Directory.CreateDirectory(tb_LogLocation.Text);
            }
            Application.DoEvents();
            fbd.Dispose();
        }
        private void btn_ClearGrid_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
        }
        private void lb_Documentation_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (File.Exists(Environment.CurrentDirectory + "\\Documentation\\" + lb_Documentation.Text.ToString()))
            {
                textBox1.Lines = File.ReadAllLines(Environment.CurrentDirectory + "\\Documentation\\" + lb_Documentation.Text.ToString());
            }
            if (File.Exists(Environment.CurrentDirectory + "\\Configs\\" + lb_Documentation.Text.ToString()))
            {
                textBox1.Lines = File.ReadAllLines(Environment.CurrentDirectory + "\\Configs\\" + lb_Documentation.Text.ToString());
            }
        }
        private void btn_Save_Click(object sender, EventArgs e)
        {
            if (File.Exists(Environment.CurrentDirectory + "\\Documentation\\" + lb_Documentation.Text.ToString()))
            {
                File.WriteAllLines(Environment.CurrentDirectory + "\\Documentation\\" + lb_Documentation.Text.ToString(), textBox1.Lines);
            }
            if (File.Exists(Environment.CurrentDirectory + "\\Configs\\" + lb_Documentation.Text.ToString()))
            {
                File.WriteAllLines(Environment.CurrentDirectory + "\\Configs\\" + lb_Documentation.Text.ToString(), textBox1.Lines);
            }
        }
    }
}
