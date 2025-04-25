using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BibleCompiler2
{
    internal class OpenFileException
    {
        public void err(string fileName)
        {
            MessageBox.Show($"Your \"{fileName}\" Document is Open,\nPlease Save And Close Your Document", "Open Document Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            string processName = "winWord";
            Process[] processList = Process.GetProcesses();
            foreach (Process theProcess in processList)
            {
                if (theProcess.ProcessName.Contains(processName))
                {
                    theProcess.Kill();
                }
            }
        }

    }
}
