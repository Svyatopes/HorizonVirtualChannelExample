using System;
using System.Threading;
using System.Windows.Forms;
// ReSharper disable LocalizableElement

namespace VMware.Horizon.VirtualChannel.Client
{
    public partial class FrmDetails : Form
    {
        public FrmDetails()
        {
            InitializeComponent();
        }

        private void frmDetails_Load(object sender, EventArgs e)
        {
            Runtime.HorizonMonitor = new HorizonMonitor();
            Runtime.HorizonMonitor.ThreadMessage += HorizonMontor_ThreadMessage;
            Runtime.HorizonMonitor.ThreadException += HorizonMonitor_ThreadException;
            if (Runtime.HorizonMonitor.Initialise())
            {
                var MonitorThread = new Thread(Runtime.HorizonMonitor.Start);
                MonitorThread.Start();
            }
            else
            {
                MessageBox.Show("Failed to register horizon client", "Failed to register Horizon client",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void FollowListBoxTail()
        {
            if (lbDebug.InvokeRequired)
            {
                lbDebug.Invoke((Action)FollowListBoxTail);
            }
            else
            {
                lbDebug.SelectedIndex = lbDebug.Items.Count - 1;
            }
        }

        private void HorizonMonitor_ThreadException(Exception ex)
        {
            if (lbDebug.InvokeRequired)
            {
                lbDebug.Invoke((Action<Exception>)HorizonMonitor_ThreadException, ex);
            }
            else
            {
                lbDebug.Items.Add(ex.ToString());
                FollowListBoxTail();
            }
        }

        private void HorizonMontor_ThreadMessage(int severity, string message)
        {
            var Entry = $"Sev: {severity} - Message: {message}";
            if (lbDebug.InvokeRequired)
            {
                lbDebug.Invoke((Action<int, string>)HorizonMontor_ThreadMessage, severity, message);
            }
            else
            {
                lbDebug.Items.Add(Entry);
                while (lbDebug.Items.Count > 100) lbDebug.Items.RemoveAt(0);
                FollowListBoxTail();
            }
        }

        private void niClient_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Visible = true;
        }

        private void tsmiExit_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void frmDetails_FormClosing(object sender, FormClosingEventArgs e)
        {
            Runtime.HorizonMonitor.IsClosing = true;
        }

        private void frmDetails_Shown(object sender, EventArgs e)
        {
            Visible = false;
        }
    }
}