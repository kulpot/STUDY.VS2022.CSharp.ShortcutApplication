using IWshRuntimeLibrary;
using System;
using System.Windows.Forms;
using System.Reflection;

namespace ShortcutApplication
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            string folder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            CreateShortcutToCurrentAssembly(folder);
        }
        
        private void CreateShortcutToCurrentAssembly(string saveDir)
        {
            //WshShellClass wshShell = new WshShellClass();
            WshShell wshShell = new WshShell();
            string fileName = saveDir + "\\" + ProductName + ".lnk";
            IWshShortcut shortcut = (IWshShortcut)wshShell.CreateShortcut(fileName);
            shortcut.TargetPath = Application.ExecutablePath;
            shortcut.Save();
        }
    }
}
