using MFilesAPI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BSVI.FileCheck
{
    public static class VaultConnect
    {
        private static void splitDomain(ref string windowUser, ref string windowDomain, string vaultLogin)
        {
            char[] delimiterChars = { '\\' };
            string[] splituser = vaultLogin.Split(delimiterChars);
            if (splituser.Length > 1)
            {
                windowUser = splituser[1];
                windowDomain = splituser[0];
            }
            else
            {
                windowUser = vaultLogin;
                windowDomain = "";
            }
        }

        public static Vault ConnectServer(
       ref string ErrMsg,
       string vautGUId,
       string vaultLogin, // baymain/jutta.ipsen --- split domain
       string vaultPWD,
       bool isWindowsUser,
       string VaultNetworkAddress,
       string VaultEndpoint,
       string VaultProtocol,
       string windowUser = "",
       string windowDomain = ""
       )
        {
            Vault _vault = null;
            try
            {
                MFilesServerApplication _server = new MFilesServerApplication();
                //MFilesClientApplication _client = new MFilesClientApplication();
                splitDomain(ref windowUser, ref windowDomain, vaultLogin);
                if (VaultNetworkAddress.Equals("localhost"))
                {
                    if (isWindowsUser == true)
                    {
                        if (windowDomain.Equals(""))
                        {
                            _server.Connect(MFAuthType.MFAuthTypeSpecificWindowsUser, windowUser, vaultPWD);
                        }
                        else
                        {
                            _server.Connect(MFAuthType.MFAuthTypeSpecificWindowsUser, windowUser, vaultPWD, windowDomain);
                        }
                    }
                    else
                    {
                        _server.Connect(MFAuthType.MFAuthTypeSpecificMFilesUser, vaultLogin, vaultPWD);
                    }
                }
                else// network
                {

                    string Protocol;
                    if (String.IsNullOrEmpty(VaultProtocol))
                    {
                        Protocol = "ncacn_ip_tcp"; //(TCP/IP protocol)
                    }
                    else
                    {
                        Protocol = VaultProtocol;
                    }
                    if (isWindowsUser == true)
                    {
                        if (windowDomain.Equals(""))
                        {
                            _server.Connect(MFAuthType.MFAuthTypeSpecificWindowsUser, windowUser, vaultPWD, Type.Missing,
                                Protocol, VaultNetworkAddress, VaultEndpoint);
                        }
                        else
                        {
                            _server.Connect(MFAuthType.MFAuthTypeSpecificWindowsUser, windowUser, vaultPWD, windowDomain,
                             Protocol, VaultNetworkAddress, VaultEndpoint);
                        }
                    }
                    else
                    {
                        _server.Connect(MFAuthType.MFAuthTypeSpecificMFilesUser, vaultLogin, vaultPWD, Type.Missing,
                            Protocol, VaultNetworkAddress, VaultEndpoint);
                    }
                }


                _vault = _server.LogInToVault(vautGUId);



                //var vaultConnection = _client.GetVaultConnection(vaultName);
                //IntPtr hwnd = Process.GetCurrentProcess().MainWindowHandle;
                //_vault = vaultConnection.BindToVault(hwnd, true, true);
                return _vault;
            }
            catch (Exception e)
            {
                ErrMsg = e.Message;
                return _vault;
            }
        }
    }
}
