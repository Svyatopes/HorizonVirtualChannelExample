using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Timers;
using Newtonsoft.Json;
using VMware.Horizon.VirtualChannel.RDPVCBridgeInterop;
using static VMware.Horizon.VirtualChannel.PipeMessages.V1;

namespace VMware.Horizon.VirtualChannel.AgentAPI
{
    public class VirtualChannelAgent
    {
        public delegate void ChannelConnectedHandler(bool connected);

        public delegate void SyncLocalVolumeHandler(VolumeStatus vs);

        public delegate void ThreadExceptionHandler(Exception ex);

        public delegate void ThreadMessageCallback(int severity, string message);

        public bool Connected;

        private bool _hasRequestedLocalAudio;

        public bool IsClosing = false;
        private Timer _pulseTimer;

        public VirtualChannelAgent(string channelName)
        {
            Lock = new object();
            var sid = Process.GetCurrentProcess().SessionId;
            Handle = RdpvcBridge.VDP_VirtualChannelOpen(VirtualChannelStructures.WTS_CURRENT_SERVER_HANDLE, sid,
                channelName);
            if (Handle == IntPtr.Zero)
            {
                var er = Marshal.GetLastWin32Error();
                throw new Exception("Could not Open the virtual Channel: " + er);
            }
        }

        public IntPtr Handle { get; set; }
        public object Lock { get; set; }

        public event ThreadMessageCallback LogMessage;

        public event ChannelConnectedHandler ChannelConnectionChange;

#pragma warning disable CS0067
        public event ThreadExceptionHandler ObjectException;
#pragma warning restore CS0067

        public event SyncLocalVolumeHandler SyncLocalVolume;

        ~VirtualChannelAgent()
        {
            if (Handle != IntPtr.Zero)
            {
                try
                {
                    RdpvcBridge.VDP_VirtualChannelClose(Handle);
                }
                catch
                {
                    // ignored
                }

                try
                {
                    _pulseTimer.Stop();
                }
                catch
                {
                    // ignored
                }

                try
                {
                    _pulseTimer.Dispose();
                }

                catch
                {
                    // ignored
                }
            }
        }

        private void InitializePulseTimer()
        {
            _pulseTimer = new Timer
            {
                Enabled = true,
                Interval = 5000
            };
            _pulseTimer.Start();
        }

        public void Destroy()
        {
            LogMessage?.Invoke(3, "Closing object.");
            RdpvcBridge.VDP_VirtualChannelClose(Handle);
            Handle = IntPtr.Zero;
            ChannelConnectionChange?.Invoke(false);
            _pulseTimer.Stop();
            _pulseTimer.Dispose();
            Connected = false;
        }

        private void ChangeConnectivity(bool connected)
        {
            if (Connected != connected)
            {
                Connected = connected;
                ChannelConnectionChange?.Invoke(Connected);
            }
        }

        public async Task<bool> Open()
        {
            InitializePulseTimer();
            _pulseTimer.Elapsed += PulseTimer_Elapsed;

            var ChannelResponse = await Probe();
            if (ChannelResponse.Successful)
            {
                return true;
            }

            return false;
        }

        private async void PulseTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            if ((await Probe()).Successful)
            {
                if (!_hasRequestedLocalAudio)
                {
                    var LocalVolume = await GetClientVolume();
                    SyncLocalVolume?.Invoke(LocalVolume);
                    _hasRequestedLocalAudio = true;
                }
            }
        }

        private async Task<object> SendMessage(ChannelCommand messageObject, Type returnType)
        {
            return await Task.Run(() =>
            {
                LogMessage?.Invoke(3, "send requested, awaiting lock.");

                lock (Lock)
                {
                    LogMessage?.Invoke(3, "send requested, lock received.");

                    var written = 0;
                    var serialisedMessage = JsonConvert.SerializeObject(messageObject);
                    LogMessage?.Invoke(3, $"Sending Message : {serialisedMessage}");
                    var msg = BinaryConverters.StringToBinary(serialisedMessage);
                    var SendResult = RdpvcBridge.VDP_VirtualChannelWrite(Handle, msg, msg.Length, ref written);
                    LogMessage?.Invoke(3,
                        string.Format("Sending Message result: {0} - Written: {1}", SendResult, written));
                    if (!SendResult)
                    {
                        LogMessage?.Invoke(2, "Sending the command was not succesful");
                        ChangeConnectivity(false);
                        return null;
                    }

                    var buffer = new byte[10240];
                    var actualRead = 0;

                    var ReceiveResult =
                        RdpvcBridge.VDP_VirtualChannelRead(Handle, 5000, buffer, buffer.Length, ref actualRead);
                    LogMessage?.Invoke(3,
                        string.Format("VDP_VirtualChannelRead result: {0} - ActualRead: {1}", ReceiveResult,
                            actualRead));
                    if (!ReceiveResult)
                    {
                        ChangeConnectivity(false);
                        LogMessage?.Invoke(3, "Did not receive a response in a timely fashion or we received an error");
                        return null;
                    }

                    var receivedContents = new byte[actualRead];
                    Buffer.BlockCopy(buffer, 0, receivedContents, 0, actualRead);
                    var serialisedResponse = BinaryConverters.BinaryToString(receivedContents);
                    LogMessage?.Invoke(3, string.Format("Received: {0}", serialisedResponse));
                    return JsonConvert.DeserializeObject(serialisedResponse, returnType, (JsonSerializerSettings)null);
                }
            });
        }

        public async Task<ChannelResponse> Probe()
        {
            try
            {
                var ProbeResponse =
                    await SendMessage(new ChannelCommand(CommandType.Probe, null), typeof(ChannelResponse));
                if (ProbeResponse != null)
                {
                    var Response = (ChannelResponse)ProbeResponse;
                    if (Response.Successful)
                    {
                        ChangeConnectivity(true);
                    }
                    else
                    {
                        ChangeConnectivity(false);
                    }

                    return Response;
                }

                ChangeConnectivity(false);
                LogMessage(1, "Receive Failed during probe");
                return new ChannelResponse { Successful = false, Details = "Receive Failed" };
            }
            catch (Exception ex)
            {
                LogMessage?.Invoke(1, string.Format("Exception trapped in Probe: {0}", ex));
                return new ChannelResponse
                {
                    Successful = false,
                    Details = ex.ToString()
                };
            }
        }

        public async Task<ChannelResponse> SetVolume(VolumeStatus sv)
        {
            if (Connected)
            {
                LogMessage?.Invoke(3, "SetVolume requested, connection open.");
                var cc = new ChannelCommand(CommandType.SetVolume, sv);
                return (ChannelResponse)await SendMessage(cc, typeof(ChannelResponse));
            }

            LogMessage?.Invoke(3, "SetVolume failed. Channel Closed.");
            ChangeConnectivity(false);
            return null;
        }

        public async Task<VolumeStatus> GetClientVolume()
        {
            if (Connected)
            {
                LogMessage?.Invoke(3, "GetVolume requested, connnection open.");
                LogMessage?.Invoke(3, "Getting Volume Status");
                return (VolumeStatus)await SendMessage(new ChannelCommand(CommandType.GetVolume, null),
                    typeof(VolumeStatus));
            }

            LogMessage?.Invoke(3, "GetVolume failed. Channel Closed.");
            ChangeConnectivity(false);
            return null;
        }
    }
}