using System;
using Microsoft.Win32;
using NLog;

namespace VMware.Horizon.VirtualChannel.RegistryHelpers
{
    public class ClientHelpers
    {
        public static string ClientPath = @"SOFTWARE\VMware, Inc.\VMware VDM\Client";

        public static bool IsAgentInstalled()
        {
            try
            {
                using (var MachineHive = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32))
                {
                    using (var ClientKey = MachineHive.OpenSubKey(ClientPath))
                    {
                        if (ClientKey != null)
                        {
                            var ClientVersion = ClientKey.GetValue("Version", null);
                            if (ClientVersion != null)
                            {
                                LogManager.GetCurrentClassLogger().Info("Client Version: {0}", ClientVersion);
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