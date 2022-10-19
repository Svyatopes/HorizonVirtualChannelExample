using System;
using System.Windows.Forms;
using VMware.Horizon.VirtualChannel.RegistryHelpers;
// ReSharper disable LocalizableElement

namespace VMware.Horizon.VirtualChannel.Client
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        private static void Main()
        {
            if (!ClientHelpers.IsAgentInstalled())
            {
                MessageBox.Show("This app cannot function without a horizon client", "Horizon Client Missing",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new FrmDetails());
            }
        }
    }
}