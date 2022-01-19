using System;
using System.Diagnostics;
using System.Net;
using System.Net.NetworkInformation;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Principal;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace COPS
{
    static class Global
    {
        public static string strFullPathToInstaller;
    }
    static class Utility
    {
        public static bool ControlValid(TextBox name)
        {
            return name.Text.Trim().Length != 0 ? true : false;
        }
        public static bool ControlValid(ComboBox name)
        {
            return name.Text.Trim().Length != 0 ? true : false;
        }
    }
    public class DIO
    {
        public double CalculateFolderSize(string strPath)
        {
            DirectoryInfo dL = new DirectoryInfo(strPath);
            FileInfo[] fL = dL.GetFiles("*.*", SearchOption.AllDirectories);
            double Lfoldersize = 0;
            for (int i = 0; i < fL.Count(); i++)
            {
                Lfoldersize += fL[i].Length;
            }
            return Lfoldersize;
        }
        public void CopyFilesRecursively(string sourcePath, string targetPath)
        {
            foreach (string dirPath in Directory.GetDirectories(sourcePath, "*", SearchOption.AllDirectories))
            {
                Directory.CreateDirectory(dirPath.Replace(sourcePath, targetPath));
            }
            foreach (string newPath in Directory.GetFiles(sourcePath, "*.*", SearchOption.AllDirectories))
            {
                File.Copy(newPath, newPath.Replace(sourcePath, targetPath), true);
            }
        }
    }
    public class DeploymentResults
    {
        public string strComputer { get; set; }
        public string strIPv4Address { get; set; }
        public string strPingStatus { get; set; }
        public string strAdminShareStatus { get; set; }
        public string strInstallResult { get; set; }
        public DeploymentResults()
        {
            strComputer = null;
            strIPv4Address = null;
            strPingStatus = null;
            strAdminShareStatus = null;
            strInstallResult = null;
        }
    }
    public class PreDeploymentChecks 
    {
        public List<bool> lstChecklist { get; set; }
        public bool CyberOrderCheck { get; set; } 
        public bool CyberOrderIDCheck { get; set; }
        public bool InstallerCheck { get; set; }
        public bool PresetsCheck { get; set; }
        public bool LocalDIRCheck { get; set; }
        public bool RemoteDIRCheck { get; set; }
        public bool ParametersCheck { get; set; }
        public bool ComputerListCheck { get; set; }
        public PreDeploymentChecks() 
        { 
            CyberOrderCheck = false; 
            CyberOrderIDCheck = false;
            InstallerCheck = false;
            PresetsCheck = false;
            LocalDIRCheck = false;
            RemoteDIRCheck = false;
            ParametersCheck = false;
            ComputerListCheck = false;
        } 
        public bool RunChecklist()
        {
            lstChecklist = new List<bool>();
            lstChecklist.Add(CyberOrderCheck);
            lstChecklist.Add(CyberOrderIDCheck);
            lstChecklist.Add(InstallerCheck);
            lstChecklist.Add(PresetsCheck);
            lstChecklist.Add(LocalDIRCheck);
            lstChecklist.Add(RemoteDIRCheck);
            lstChecklist.Add(ParametersCheck);
            lstChecklist.Add(ComputerListCheck);
            return !lstChecklist.Contains(false);
        }
    }
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
