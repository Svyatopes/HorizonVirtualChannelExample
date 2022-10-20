using System;
using System.Collections.Generic;
using System.Threading;
using NAudio.CoreAudioApi;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using Newtonsoft.Json.Linq;
using VMware.Horizon.Interop;
using VMware.Horizon.VirtualChannel.RDPVCBridgeInterop;
using static VMware.Horizon.VirtualChannel.PipeMessages.V1;

namespace VMware.Horizon.VirtualChannel.Client
{
    public class HorizonMonitor
    {
        public delegate void ThreadExceptionHandler(Exception ex);

        public delegate void ThreadMessageCallback(int severity, string message);

        public string BatteryStatus = "";
        private VMwareHorizonVirtualChannelEvents _channelEvents;
        private MMDevice _device;

        private MMDeviceEnumerator _en;

        public bool IsClosing;

        //private IVMwareHorizonClientVChan VMwareHorizonVirtualChannelAPI = null;

        private IVMwareHorizonClient4 _vmhc;

        private void SetVolumeStatus(VolumeStatus sv)
        {
            using (var en = new MMDeviceEnumerator())
            {
                var device = en.GetDefaultAudioEndpoint(DataFlow.Render,
                    Role.Console);
                device.AudioEndpointVolume.Mute = sv.Muted;
                device.AudioEndpointVolume.MasterVolumeLevelScalar = sv.VolumeLevel;
            }
        }

        private VolumeStatus GetVolume()
        {
            using (var en = new MMDeviceEnumerator())
            {
                var device = en.GetDefaultAudioEndpoint(DataFlow.Render,
                    Role.Console);
                return new VolumeStatus(device.AudioEndpointVolume.Mute,
                    device.AudioEndpointVolume.MasterVolumeLevelScalar);
            }
        }

        private void AudioEndpointVolume_OnVolumeNotification(AudioVolumeNotificationData data)
        {
            if (data.Muted)
            {
                ThreadMessage?.Invoke(3, "Volume Muted");
            }
            else
            {
                ThreadMessage?.Invoke(3, "Volume changed to: " + data.MasterVolume);
            }
        }


        public event ThreadMessageCallback ThreadMessage;
        public event ThreadExceptionHandler ThreadException;

        public bool Initialise()
        {
            // Open Audio Callbacks
             
            _en = new MMDeviceEnumerator();
            _device = _en.GetDefaultAudioEndpoint(DataFlow.Render, Role.Console);
            _device.AudioEndpointVolume.OnVolumeNotification += AudioEndpointVolume_OnVolumeNotification;
            GC.SuppressFinalize(_en);
            GC.SuppressFinalize(_device);
            ThreadMessage?.Invoke(3, "Opened Audio API");


            // Open Horizon Client Listener
            _vmhc = (IVMwareHorizonClient4)new VMwareHorizonClient();
            IVMwareHorizonClientEvents5
                HorizonEvents = new VMwareHorizonClientEvents(this);
            _vmhc.Advise2(HorizonEvents, VmwHorizonClientAdviseFlags.VmwHorizonClientAdvise_DispatchCallbacksOnUIThread);
            GC.SuppressFinalize(_vmhc);
            ThreadMessage?.Invoke(3, "Opened Horizon API");

            // Register Virtual Channel Callback
            _channelEvents = new VMwareHorizonVirtualChannelEvents(this);
            var Channels = new VMwareHorizonClientChannelDefinition[1];
            Channels[0] = new VMwareHorizonClientChannelDefinition("VVCAM", 0);
            _vmhc.RegisterVirtualChannelConsumer2(Channels, _channelEvents, out var apiObject);
            _channelEvents.HorizonClientVirtualChannel = (IVMwareHorizonClientVChan)apiObject;
            GC.SuppressFinalize(_channelEvents);
            ThreadMessage.Invoke(3, "Opened Virtual Channel Listener");
            return true;
        }


        public void Start()
        {
            try
            {
                while (!IsClosing) Thread.Sleep(500);

                _en.Dispose();
                _device.Dispose();

                GC.ReRegisterForFinalize(_channelEvents);
                GC.ReRegisterForFinalize(_en);
                GC.ReRegisterForFinalize(_device);
                GC.ReRegisterForFinalize(_vmhc);
            }
            catch (Exception ex)
            {
                ThreadMessage?.Invoke(1,
                    string.Format("The Horizon Monitor thread reported a fatal Exception: {0}", ex));
                ThreadException?.Invoke(ex);
            }
        }

        public void Close()
        {
            IsClosing = true;
        }


        public class VMwareHorizonClientChannelDefinition : IVMwareHorizonClientChannelDef
        {
            public VMwareHorizonClientChannelDefinition(string name, uint options)
            {
                this.name = name;
                this.options = options;
            }

            public string name { get; }

            public uint options { get; }
        }

        public class VMwareHorizonVirtualChannelEvents : IVMwareHorizonClientVChanEvents
        {
            private readonly HorizonMonitor _callbackObject;
            public IVMwareHorizonClientVChan HorizonClientVirtualChannel;

            private uint _mChannelHandle;
            private byte[] _mClientPingFragment = { 0x50 /* 'P' */, 0x6F /* 'o' */, 0x6E /* 'n' */, 0x67 /* 'g' */ };
            private int _mPingTestCurLen;
            private byte[] _mPingTestMsg;
            private byte[] _mServerPingFragment = { 0x50 /* 'P' */, 0x69 /* 'i' */, 0x6E /* 'n' */, 0x67 /* 'g' */ };

            public VMwareHorizonVirtualChannelEvents(HorizonMonitor callback)
            {
                _callbackObject = callback;
            }

            public void ConnectEventProc(uint serverId, string sessionToken, uint eventType, Array eventData)
            {
                var currentEventType =
                    (VirtualChannelStructures.ChannelEvents)eventType;
                _callbackObject.ThreadMessage?.Invoke(3, "ConnectEventProc() called: " + currentEventType);
                //  SharedObjects.hvm.ThreadMessage?.Invoke(3, "ConnectEventProc() called ");

                if (eventType == (uint)VirtualChannelStructures.ChannelEvents.Connected)

                {
                    try
                    {
                        HorizonClientVirtualChannel.VirtualChannelOpen(serverId, sessionToken, "VVCAM",
                            out _mChannelHandle);
                        _callbackObject.ThreadMessage?.Invoke(3, "!! VirtualChannelOpen() succeeded");
                    }
                    catch (Exception ex)
                    {
                        _callbackObject.ThreadMessage?.Invoke(3,
                            string.Format("VirtualChannelOpen() failed: {0}", ex));
                        _mChannelHandle = 0;
                    }
                }
            }

            public void InitEventProc(uint serverId, string sessionToken, uint rc)
            {
                _callbackObject.ThreadMessage?.Invoke(3, "InitEventProc()");
            }

            public void ReadEventProc(uint serverId, string sessionToken, uint channelHandle, uint eventType,
                Array eventData, uint totalLength, uint dataFlags)
            {
                var currentEventType =
                    (VirtualChannelStructures.ChannelEvents)eventType;
                var cf =
                    (VirtualChannelStructures.ChannelFlags)dataFlags;
                _callbackObject.ThreadMessage?.Invoke(3,
                    "ReadEventProc(): " + currentEventType + " - Flags: " + cf + " - Length: " +
                    totalLength);

                var isFirst = (dataFlags & (uint)VirtualChannelStructures.ChannelFlags.First) != 0;
                var isLast = (dataFlags & (uint)VirtualChannelStructures.ChannelFlags.Last) != 0;

                if (isFirst)
                {
                    _mPingTestMsg = new byte[totalLength];
                    _mPingTestCurLen = 0;
                }

                eventData.CopyTo(_mPingTestMsg, _mPingTestCurLen);
                _mPingTestCurLen += eventData.Length;

                if (isLast)
                {
                    if (totalLength != _mPingTestMsg.Length)
                    {
                        _callbackObject.ThreadMessage?.Invoke(3,
                            "Received {mPingTestMsg.Length} bytes but expected {totalLength} bytes!");
                    }

                    var message = BinaryConverters.BinaryToString(_mPingTestMsg);
                    var cc = JsonConvert.DeserializeObject<ChannelCommand>(message);
                    _callbackObject.ThreadMessage?.Invoke(3,
                        "Received: " + cc.CommandType + " = " +
                        BinaryConverters.BinaryToString(_mPingTestMsg));

                    try
                    {
                        switch (cc.CommandType)
                        {
                            case CommandType.SetVolume:
                                var jo = (JObject)cc.CommandParameters;
                                var sv = jo.ToObject<VolumeStatus>();
                                _callbackObject.SetVolumeStatus(sv);
                                HorizonClientVirtualChannel.VirtualChannelWrite(serverId, sessionToken, channelHandle,
                                    BinaryConverters.StringToBinary(
                                        JsonConvert.SerializeObject(new ChannelResponse())));
                                break;
                            case CommandType.Probe:
                                HorizonClientVirtualChannel.VirtualChannelWrite(serverId, sessionToken, channelHandle,
                                    BinaryConverters.StringToBinary(
                                        JsonConvert.SerializeObject(new ChannelResponse())));
                                break;
                            case CommandType.GetVolume:
                                HorizonClientVirtualChannel.VirtualChannelWrite(serverId, sessionToken, channelHandle,
                                    BinaryConverters.StringToBinary(
                                        JsonConvert.SerializeObject(_callbackObject.GetVolume())));
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        _callbackObject.ThreadMessage?.Invoke(3,
                            string.Format("VirtualChannelWrite failed: {0}", ex));
                    }
                }
            }
        }

        public class VMwareHorizonClientEvents : IVMwareHorizonClientEvents5
        {
            private readonly HorizonMonitor _callbackObject;

            public VMwareHorizonClientEvents(HorizonMonitor callback)
            {
                _callbackObject = callback;
            }

            public void OnStarted()
            {
                DispatchMessage(3, "Started Called");
            }

            public void OnExit()
            {
                DispatchMessage(3, "Exit Called");
            }

            public void OnConnecting(object serverInfo)
            {
                var Info = (IVMwareHorizonClientServerInfo)serverInfo;
                DispatchMessage(3, string.Format("Connecting, Server Address: {0}, ID: {1}, Type:{2} ",
                    Info.serverAddress, Info.serverId, Info.serverType.ToString()));
            }

            public void OnConnectFailed(uint serverId, string errorMessage)
            {
                DispatchMessage(3, string.Format("Connect Failed, Server ID: {0}, Message: {1}",
                    serverId, errorMessage));
            }

            public void OnAuthenticationRequested(uint serverId, VmwHorizonClientAuthType authType)
            {
                DispatchMessage(3, string.Format("Authentication Requested, Server ID: {0}, AuthType: {1}",
                    serverId, authType.ToString()));
            }

            public void OnAuthenticating(uint serverId, VmwHorizonClientAuthType authType, string user)
            {
                DispatchMessage(3, string.Format("Authenticating, Server ID: {0}, AuthType: {1}, User: {2}",
                    serverId, authType.ToString(), user));
            }

            public void OnAuthenticationDeclined(uint serverId, VmwHorizonClientAuthType authType)
            {
                DispatchMessage(3, string.Format("Authentication Declined, Server ID: {0}, AuthType: {1}",
                    serverId, authType.ToString()));
            }

            public void OnAuthenticationFailed(uint serverId, VmwHorizonClientAuthType authType, string errorMessage,
                int retryAllowed)
            {
                DispatchMessage(3,
                    string.Format(
                        "Authentication Failed, Server ID: {0}, AuthType: {1}, Error: {2}, retry allowed?: {3}",
                        serverId, authType.ToString(), errorMessage, retryAllowed));
            }

            public void OnLoggedIn(uint serverId)
            {
                DispatchMessage(3, string.Format("Logged In, Server ID: {0}", serverId));
            }

            public void OnDisconnected(uint serverId)
            {
                DispatchMessage(3, string.Format("Disconnected, Server ID: {0}", serverId));
            }

            public void OnReceivedLaunchItems(uint serverId, Array launchItems)
            {
                DispatchMessage(3, string.Format("Received Launch Items, Server ID: {0}, Item Count: {1}", serverId,
                    launchItems.Length));
                var Items = Helpers.GetLaunchItems(launchItems);
                foreach (var item in Items)
                {
                    DispatchMessage(3,
                        string.Format("Launch Item: Server ID: {0}, Name: {1}, Type: {2}, ID: {3}", serverId, item.Name,
                            item.Type.ToString(), item.Id));
                }
            }

            public void OnLaunchingItem(uint serverId, VmwHorizonLaunchItemType type, string launchItemId,
                VmwHorizonClientProtocol protocol)
            {
                DispatchMessage(3,
                    string.Format("Launching Item, Server ID: {0}, type: {1}, Item ID: {2}, Protocol: {3}", serverId,
                        type.ToString(), launchItemId, protocol.ToString()));
            }

            public void OnItemLaunchSucceeded(uint serverId, VmwHorizonLaunchItemType type, string launchItemId)
            {
                DispatchMessage(3,
                    string.Format("Launch Item Succeeded, Server ID: {0}, Type: {1}, ID: {2}", serverId,
                        type.ToString(), launchItemId));
            }

            public void OnItemLaunchFailed(uint serverId, VmwHorizonLaunchItemType type, string launchItemId,
                string errorMessage)
            {
                DispatchMessage(3, string.Format("Launch Item Succeeded, Server ID: {0}, type: {1}, Item ID: {2}",
                    serverId,
                    type.ToString(), launchItemId));
            }

            public void OnNewProtocolSessionCreated(uint serverId, string sessionToken,
                VmwHorizonClientProtocol protocol, VmwHorizonClientSessionType type, string clientId)
            {
                DispatchMessage(3,
                    string.Format(
                        "New Protocol Session Created, Server ID: {0}, Token: {1}, Protocol: {2}, Type: {3}, ClientID: {4}",
                        serverId, sessionToken, protocol.ToString(), type.ToString(), clientId));
            }

            public void OnProtocolSessionDisconnected(uint serverId, string sessionToken, uint connectionFailed,
                string errorMessage)
            {
                DispatchMessage(3, string.Format("" +
                                                 "Protocol Session Disconnected, Server ID: {0}, Token: {1}, ConnectFailed: {2}, Error: {3}",
                    serverId, sessionToken, connectionFailed, errorMessage));
            }

            public void OnSeamlessWindowsModeChanged(uint serverId, string sessionToken, uint enabled)
            {
                DispatchMessage(3,
                    string.Format("Seamless Window Mode Changed, Server ID: {0}, Token: {1}, Enabled: {2}",
                        serverId, sessionToken, enabled));
            }

            public void OnSeamlessWindowAdded(uint serverId, string sessionToken, string windowPath,
                string entitlementId, int windowId, long windowHandle, VmwHorizonClientSeamlessWindowType type)
            {
                DispatchMessage(3, string.Format(
                    "Seamless Window Added, Server ID: {0}, Token: {1}, WindowPath: {2}, EntitlementID: {3}, WindowID: {4}, WindowHandle: {5}, Type: {6}",
                    serverId, sessionToken, windowPath, entitlementId, windowId, windowHandle, type.ToString()));
            }

            public void OnSeamlessWindowRemoved(uint serverId, string sessionToken, int windowId)
            {
                DispatchMessage(3, string.Format(
                    "Seamless Window Removed, Server ID: {0}, Token: {1}, WindowID: {2}",
                    serverId, sessionToken, windowId));
            }

            public void OnUSBInitializeComplete(uint serverId, string sessionToken)
            {
                DispatchMessage(3, string.Format(
                    "USB Initialize Complete, Server ID: {0}, Token: {1}",
                    serverId, sessionToken));
            }

            public void OnConnectUSBDeviceComplete(uint serverId, string sessionToken, uint isConnected)
            {
                DispatchMessage(3, string.Format(
                    "Connect USB Device Complete, Server ID: {0}, Token: {1}, IsConnected: {2}",
                    serverId, sessionToken, isConnected));
            }

            public void OnUSBDeviceError(uint serverId, string sessionToken, string errorMessage)
            {
                DispatchMessage(3, string.Format(
                    "Connect USB Device Error, Server ID: {0}, Token: {1}, Error: {2}",
                    serverId, sessionToken, errorMessage));
            }

            public void OnAddSharedFolderComplete(uint serverId, string fullPath, uint succeeded, string errorMessage)
            {
                DispatchMessage(3, string.Format(
                    "Add Shared Folder Complete, Server ID: {0}, FullPath: {1}, Succeeded: {2}, Error: {3}",
                    serverId, fullPath, succeeded, errorMessage));
            }

            public void OnRemoveSharedFolderComplete(uint serverId, string fullPath, uint succeeded,
                string errorMessage)
            {
                DispatchMessage(3, string.Format(
                    "Remove Shared Folder Complete, Server ID: {0}, FullPath: {1}, Succeeded: {2}, Error: {3}",
                    serverId, fullPath, succeeded, errorMessage));
            }

            public void OnFolderCanBeShared(uint serverId, string sessionToken, uint canShare)
            {
                DispatchMessage(3, string.Format(
                    "Folder Can Be Shared, Server ID: {0}, Token: {1}, canShare: {2}",
                    serverId, sessionToken, canShare));
            }

            public void OnCDRForcedByAgent(uint serverId, string sessionToken, uint forcedByAgent)
            {
                DispatchMessage(3, string.Format(
                    "CDR Forced By Agent, Server ID: {0}, Token: {1}, Forced: {2}",
                    serverId, sessionToken, forcedByAgent));
            }

            public void OnItemLaunchSucceeded2(uint serverId, VmwHorizonLaunchItemType type, string launchItemId,
                string sessionToken)
            {
                DispatchMessage(3,
                    string.Format("Item Launch Succeeded(2), Server ID: {0}, Type: {1}, ID: {2}, token: {3}", serverId,
                        type.ToString(), launchItemId, sessionToken));
            }

            public void OnReceivedLaunchItems2(uint serverId, Array launchItems)
            {
                DispatchMessage(3,
                    string.Format("Received Launch Items2, Server ID: {0}, Item Count: {1}", serverId,
                        launchItems.Length));
                var Items = Helpers.GetLaunchItems2(launchItems);
                foreach (var item in Items)
                {
                    DispatchMessage(3,
                        string.Format("Launch Item: Server ID: {0}, Name: {1}, Type: {2}, ID: {3}, Remotable: {4}",
                            serverId, item.Name, item.Type.ToString(), item.Id, item.HasRemotableAssets));
                }
            }

            private void DispatchMessage(int severity, string message)
            {
                _callbackObject.ThreadMessage?.Invoke(severity, message);
            }

            private string SerialiseObject(object value) =>
                JsonConvert.SerializeObject(value, Formatting.Indented);

            public class Helpers
            {
                [Flags]
                public enum LaunchItemType
                {
                    VmwHorizonLaunchItem_HorizonDesktop = 0,
                    VmwHorizonLaunchItem_HorizonApp = 1,
                    VmwHorizonLaunchItem_XenApp = 2,
                    VmwHorizonLaunchItem_SaaSApp = 3,
                    VmwHorizonLaunchItem_HorizonAppSession = 4,
                    VmwHorizonLaunchItem_DesktopShadowSession = 5,
                    VmwHorizonLaunchItem_AppShadowSession = 6
                }

                [Flags]
                public enum SupportedProtocols
                {
                    VmwHorizonClientProtocol_Default = 0,
                    VmwHorizonClientProtocol_RDP = 1,
                    VmwHorizonClientProtocol_PCoIP = 2,
                    VmwHorizonClientProtocol_Blast = 4
                }


                public static List<LaunchItem> GetLaunchItems(Array ItemList)
                {
                    var returnList = new List<LaunchItem>();
                    foreach (var item in ItemList)
                    {
                        returnList.Add(new LaunchItem((IVMwareHorizonClientLaunchItemInfo)item));
                    }

                    return returnList;
                }

                public static List<LaunchItem2> GetLaunchItems2(Array ItemList)
                {
                    var returnList = new List<LaunchItem2>();
                    foreach (var item in ItemList)
                    {
                        returnList.Add(new LaunchItem2((IVMwareHorizonClientLaunchItemInfo2)item));
                    }

                    return returnList;
                }

                public class LaunchItem
                {
                    public LaunchItem(IVMwareHorizonClientLaunchItemInfo item)
                    {
                        Name = item.name;
                        Id = item.id;
                        Type = (LaunchItemType)item.type;
                        SupportedProtocols = (SupportedProtocols)item.supportedProtocols;
                        DefaultProtocol = item.defaultProtocol;
                    }

                    public string Name { get; set; }

                    public string Id { get; set; }

                    [JsonConverter(typeof(StringEnumConverter))]
                    public LaunchItemType Type { get; set; }

                    [JsonConverter(typeof(StringEnumConverter))]
                    public SupportedProtocols SupportedProtocols { get; set; }

                    [JsonConverter(typeof(StringEnumConverter))]
                    public VmwHorizonClientProtocol DefaultProtocol { get; set; }
                }

                public class LaunchItem2
                {
                    public LaunchItem2(IVMwareHorizonClientLaunchItemInfo2 i)
                    {
                        Name = i.name;
                        Id = i.id;
                        Type = (LaunchItemType)i.type;
                        SupportedProtocols = (SupportedProtocols)i.supportedProtocols;
                        DefaultProtocol = i.defaultProtocol;
                        HasRemotableAssets = i.hasRemotableAssets;
                    }

                    public string Name { get; set; }

                    public string Id { get; set; }

                    [JsonConverter(typeof(StringEnumConverter))]
                    public LaunchItemType Type { get; set; }

                    [JsonConverter(typeof(StringEnumConverter))]
                    public SupportedProtocols SupportedProtocols { get; set; }

                    [JsonConverter(typeof(StringEnumConverter))]
                    public VmwHorizonClientProtocol DefaultProtocol { get; set; }

                    public uint HasRemotableAssets { get; set; }
                }
            }
        }
    }
}