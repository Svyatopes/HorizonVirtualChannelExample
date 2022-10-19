using System;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;
using NAudio.CoreAudioApi;
using NLog;
using VMware.Horizon.VirtualChannel.AgentAPI;
using VMware.Horizon.VirtualChannel.RDPVCBridgeInterop;
using static VMware.Horizon.VirtualChannel.PipeMessages.V1;
// ReSharper disable LocalizableElement

namespace VMware.Horizon.VirtualChannel.Agent
{
    public partial class FrmDetails : Form
    {
        private MMDevice _device;

        private MMDeviceEnumerator _en;


        private VirtualChannelAgent _vca;


        public FrmDetails()
        {
            InitializeComponent();
        }


        private void AudioEndpointVolume_OnVolumeNotification(AudioVolumeNotificationData data)
        {
            AgentThread_ThreadMessage(3,
                $"audio Changed, Muted: {data.Muted}, Volume: {data.MasterVolume}");
            if (tsmiSyncVolume.Checked)
            {
                var sv = new VolumeStatus(data.Muted, data.MasterVolume);
                if (_vca.Connected)
                {
#pragma warning disable CS4014
                    _vca.SetVolume(sv);
#pragma warning restore CS4014
                }
            }
        }

        public void OpenAudio()
        {
            try
            {
                AgentThread_ThreadMessage(3, "Mapping audio components");
                _en = new MMDeviceEnumerator();
                _device = _en.GetDefaultAudioEndpoint(DataFlow.Render, Role.Console);
                AgentThread_ThreadMessage(3, "adding event for audio components");
                AgentThread_ThreadMessage(3, "GC KeepAlive audio components");
                GC.SuppressFinalize(_device);
                GC.SuppressFinalize(_en);
                _device.AudioEndpointVolume.OnVolumeNotification += AudioEndpointVolume_OnVolumeNotification;
            }
            catch (Exception ex)
            {
                if (_en != null)
                {
                    try
                    {
                        _en.Dispose();
                    }
                    catch
                    {
                        // ignored
                    }
                }

                if (_device != null)
                {
                    try
                    {
                        _device.Dispose();
                    }
                    catch
                    {
                        // ignored
                    }
                }

                AgentThread_ThreadMessage(1, string.Format("Failed to map Audio Device: {0}", ex));
            }
        }

        public void CloseAudio()
        {
            try
            {
                _device.AudioEndpointVolume.OnVolumeNotification -= AudioEndpointVolume_OnVolumeNotification;
                _en.Dispose();
                _device.Dispose();
                GC.ReRegisterForFinalize(_device);
                GC.ReRegisterForFinalize(_en);
            }
            catch
            {
                // ignored
            }
        }


        private void OpenAgent()
        {
            lbDetails.Items.Add("Registering Agent Thread");
            _vca = new VirtualChannelAgent("VVCAM");
            _vca.LogMessage += AgentThread_ThreadMessage;
            _vca.ObjectException += AgentThread_ThreadException;
            _vca.ChannelConnectionChange += VCA_ChannelConnectionChange;
            _vca.SyncLocalVolume += VCA_SyncLocalVolume;
            var Result = _vca.Open();
            AgentThread_ThreadMessage(3, $"Agent Open Response: {Result}");
        }

        private void VCA_SyncLocalVolume(VolumeStatus vs)
        {
            if (_en != null)
            {
                if (_device != null)
                {
                    _device.AudioEndpointVolume.MasterVolumeLevelScalar = vs.VolumeLevel;
                    _device.AudioEndpointVolume.Mute = vs.Muted;
                }
            }
        }

        private void CloseAgent()
        {
            lbDetails.Items.Add("Unregistering Agent Thread");
            if (_vca != null)
            {
                _vca.LogMessage -= AgentThread_ThreadMessage;
                _vca.ObjectException -= AgentThread_ThreadException;
                _vca.ChannelConnectionChange -= VCA_ChannelConnectionChange;
                _vca.SyncLocalVolume -= VCA_SyncLocalVolume;
                _vca.Destroy();
                _vca = null;
            }
        }

        private void VCA_ChannelConnectionChange(bool Connected)
        {
            AgentThread_ThreadMessage(3, $"Connectivity Change: {Connected}");
            niAgent.Text = $"Pipe Agent - Connected: {Connected}";
        }

        private void frmDetails_Load(object sender, EventArgs e)
        {
            if (!RdpvcBridge.VDP_IsViewSession((uint)Process.GetCurrentProcess().SessionId))
            {
                MessageBox.Show("This is not a Horizon Session, closing", "Error", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                Close();
            }
            else
            {
                SystemEvents.SessionSwitch += SystemEvents_SessionSwitch;
                OpenAudio();
                try
                {
                    OpenAgent();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format("Could not open virtual channel: {0}", ex), "Failed to start",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    Close();
                }
            }
        }

        private async void SystemEvents_SessionSwitch(object sender, SessionSwitchEventArgs e)
        {
            lbDetails.Items.Add(string.Format("Session Change: {0}", e.Reason.ToString()));
            switch (e.Reason)
            {
                case SessionSwitchReason.ConsoleDisconnect:
                case SessionSwitchReason.RemoteDisconnect:
                    CloseAgent();
                    CloseAudio();
                    break;
                case SessionSwitchReason.RemoteConnect:
                case SessionSwitchReason.ConsoleConnect:
                    var isView = RdpvcBridge.VDP_IsViewSession((uint)Process.GetCurrentProcess().SessionId);
                    lbDetails.Items.Add(string.Format("IsViewSession: {0}", isView));
                    if (isView)
                    {
                        // giving the audio component a chance to catchup
                        await Task.Run(() => { Task.Delay(3000); });
                        OpenAudio();
                        OpenAgent();
                    }

                    break;
            }
        }


        private void FollowListBoxTail()
        {
            if (lbDetails.InvokeRequired)
            {
                lbDetails.Invoke((Action)FollowListBoxTail);
            }
            else
            {
                if (tsmiFollowTail.Checked)
                {
                    lbDetails.SelectedIndex = lbDetails.Items.Count - 1;
                }
            }
        }

        private void AgentThread_ThreadException(Exception ex)
        {
            if (lbDetails.InvokeRequired)
            {
                lbDetails.Invoke((Action<Exception>)AgentThread_ThreadException, ex);
            }
            else
            {
                LogManager.GetCurrentClassLogger().Error("{0}", ex.ToString());
                lbDetails.Items.Add(ex.ToString());
                FollowListBoxTail();
            }
        }

        private void AgentThread_ThreadMessage(int severity, string message)
        {
            var Entry = string.Format("Sev: {0} - Message: {1}", severity, message);
            if (lbDetails.InvokeRequired)
            {
                lbDetails.Invoke((Action<int, string>)AgentThread_ThreadMessage, severity, message);
            }
            else
            {
                LogManager.GetCurrentClassLogger().Info("{0} - {1}", severity, message);
                lbDetails.Items.Add(Entry);
                while (lbDetails.Items.Count > 100) lbDetails.Items.RemoveAt(0);
                FollowListBoxTail();
            }
        }

        private void frmDetails_FormClosing(object sender, FormClosingEventArgs e)
        {
            CloseAgent();
            CloseAudio();
        }

        private void tsmiExit_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void niClient_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Visible = true;
        }

        private void frmDetails_Shown(object sender, EventArgs e)
        {
            Visible = false;
        }

        private void tsmiFollowTail_Click(object sender, EventArgs e)
        {
            tsmiFollowTail.Checked = !tsmiFollowTail.Checked;
        }

        private void tsmiHide_Click(object sender, EventArgs e)
        {
            Visible = false;
        }

        private void tsmiSyncVolume_Click(object sender, EventArgs e)
        {
            tsmiSyncVolume.Checked = !tsmiSyncVolume.Checked;
        }
    }
}