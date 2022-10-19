using System;
using Microsoft.Win32;
using NLog;

namespace VMware.Horizon.VirtualChannel.RegistryHelpers
{
    public class AgentHelpers
    {
        public static string AgentPath = @"SOFTWARE\VMware, Inc.\VMware VDM";

        public static bool IsAgentInstalled()
        {
            try
            {
                using (var MachineHive = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
                {
                    using (var AgentKey = MachineHive.OpenSubKey(AgentPath))
                    {
                        if (AgentKey != null)
                        {
                            var AgentVersion = AgentKey.GetValue("ProductVersion", null);
                            if (AgentVersion != null)
                            {
                                LogManager.GetCurrentClassLogger().Info("Agent Version: {0}", AgentVersion);
                                return true;
                            }
                        }

                        LogManager.GetCurrentClassLogger().Error("VMware Horizon Client not detected");
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.GetCurrentClassLogger()
                    .Error("Failed to validate Horizon Client installation: {0}", ex.ToString());
                return false;
            }
        }
    }
}