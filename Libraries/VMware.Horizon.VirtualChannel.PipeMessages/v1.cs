namespace VMware.Horizon.VirtualChannel.PipeMessages
{
    public class V1
    {
        public enum CommandType
        {
            SetVolume,
            GetSecurity,
            Probe,
            GetVolume
        }

        public class VolumeStatus
        {
            public VolumeStatus(bool muted, float volumeLevel)
            {
                Muted = muted;
                VolumeLevel = volumeLevel;
            }

            public bool Muted { get; set; }
            public float VolumeLevel { get; set; }
        }

        public class ChannelResponse
        {
            public ChannelResponse(bool success, string details)
            {
                Successful = success;
                Details = details;
            }

            public ChannelResponse()
            {
                Successful = true;
            }

            public bool Successful { get; set; }
            public string Details { get; set; }
        }

        public class ChannelCommand
        {
            public ChannelCommand(CommandType ct, object parameters)
            {
                CommandType = ct;
                CommandParameters = parameters;
            }

            public CommandType CommandType { get; set; }
            public object CommandParameters { get; set; }
        }
    }
}