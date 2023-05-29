using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace yt_DesignUI
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения..
        /// </summary>
        [STAThread]
        static void Main()
        {    
            Animator.Start();
            if (Environment.OSVersion.Version.Major >= 6) SetProcessDPIAware();

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());

            //Application.Run(new frmMain());
        }
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();
    }
}
