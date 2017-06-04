/*
MIT License

Copyright (c) 2017 Samuel Yvon

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/

using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Security.Permissions;
using System.Security.Principal;
using System.Windows.Forms;

namespace CopyPath
{
    class Program
    {
        private const string ReinstallCommand = "reinstall";
        private const string RemoveCommand = "remove";

        private const string ShellSubkeyFile = @"*\shell";
        private const string ShellSubKeyFolder = @"Directory\shell";
        private const string OperationName = "Copy full path";

        public static bool IsAdmin()
        {
            using (WindowsIdentity identity = WindowsIdentity.GetCurrent())
            {
                return (new WindowsPrincipal(identity))
                    .IsInRole(WindowsBuiltInRole.Administrator);
            }
        }


        [STAThread]
        static void Main(string[] args)
        {
            if (args.Length > 0)
            {
                string firstArg = args.First();

                if (firstArg.StartsWith("-"))
                {
                    string command = firstArg.Substring(1);

                    switch (command)
                    {
                        case ReinstallCommand:
                            if (Remove())
                            {
                                if (Install())
                                {
                                    MessageBox.Show("The copy path link has now been installed!", "Success!");
                                }
                                else
                                {
                                    ShowUnkownErrorMessage();
                                }
                            }
                            else
                            {
                                ShowUnkownErrorMessage();
                            }
                            break;
                        case RemoveCommand:
                            if (Remove())
                            {
                                MessageBox.Show("The copy path link has now been removed!", "Success!");
                            }
                            else
                            {
                                ShowUnkownErrorMessage();
                            }
                            break;
                        default:
                            MessageBox.Show("Invalid command option, please see the github page", "Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            break;
                    }
                }
                else
                {
                    Clipboard.SetText(args[0]);
                }
            }
            else
            {
                if (Install())
                {
                    MessageBox.Show("The copy path link has now been installed!", "Success!");
                }
                else
                {
                    ShowUnkownErrorMessage();
                }
            }
        }

        /// <summary>
        /// InstallReg the registry settings to add the copy path
        /// link
        /// </summary>
        private static bool Install()
        {
            if (!IsAdmin())
            {
                ShowAdminRightsRequiredMessage();
                return false;
            }

            string filePath = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName;

            if (string.IsNullOrEmpty(filePath))
            {
                ShowUnkownErrorMessage();
            }
            else
            {
                try
                {
                    bool success = AddLinkToFiles(filePath);

                    if (success)
                    {
                        return AddLinkToFolders(filePath);
                    }
                }
                catch (System.Security.SecurityException)
                {
                    ShowAdminRightsRequiredMessage();
                }
            }

            return false;
        }


        [PrincipalPermission(SecurityAction.Assert, Role = @"BUILTIN\Administrators")]
        private static bool Remove()
        {
            try
            {
                if (RemoveLinkToFiles())
                {
                    return RemoveLinkToFolders();
                }
            }
            catch (System.Security.SecurityException)
            {
                ShowAdminRightsRequiredMessage();
            }
            return false;
        }

        /// <summary>
        /// Remove the registry modifications for the individual files
        /// </summary>
        /// <returns></returns>
        [PrincipalPermission(SecurityAction.Assert, Role = @"BUILTIN\Administrators")]
        private static bool RemoveLinkToFiles()
        {
            return RemoveReg(ShellSubkeyFile);
        }

        /// <summary>
        /// Remove the registry modifications for the individual folders
        /// </summary>
        /// <returns></returns>
        [PrincipalPermission(SecurityAction.Assert, Role = @"BUILTIN\Administrators")]
        private static bool RemoveLinkToFolders()
        {
            return RemoveReg(ShellSubKeyFolder);
        }

        private static bool RemoveReg(string regKey)
        {
            using (RegistryKey shellKey = Registry.ClassesRoot.OpenSubKey(regKey, true))
            {
                if (null == shellKey)
                {
                    ShowUnkownErrorMessage("Unable to get class root reg key");
                    return false;
                }

                string[] shellKeyChildsName = shellKey.GetSubKeyNames();

                ISet<string> shellKeyChilds = new HashSet<string>(shellKeyChildsName);

                if (shellKeyChilds.Contains(OperationName))
                {
                    shellKey.DeleteSubKeyTree(OperationName, true);
                    return true;
                }

                return true;
            }
        }

        /// <summary>
        /// Add the registry modifications for the files
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        [PrincipalPermission(SecurityAction.Assert, Role = @"BUILTIN\Administrators")]
        private static bool AddLinkToFiles(string path)
        {
            return InstallReg(path, ShellSubkeyFile);
        }

        /// <summary>
        /// Add the registry modifications for the folders
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        [PrincipalPermission(SecurityAction.Assert, Role = @"BUILTIN\Administrators")]
        private static bool AddLinkToFolders(string path)
        {
            return InstallReg(path, ShellSubKeyFolder);
        }

        private static bool InstallReg(string path, string regKey)
        {
            using (RegistryKey shellKey = Registry.ClassesRoot.OpenSubKey(regKey, true))
            {
                if (null == shellKey)
                {
                    ShowUnkownErrorMessage("Unable to get class root reg key");
                    return false;
                }

                string[] shellKeyChildsName = shellKey.GetSubKeyNames();

                ISet<string> shellKeyChilds = new HashSet<string>(shellKeyChildsName);

                if (!shellKeyChilds.Contains(OperationName))
                {
                    using (RegistryKey newkey = shellKey.CreateSubKey(OperationName))
                    {
                        if (null != newkey)
                        {
                            using (RegistryKey subNewkey = newkey.CreateSubKey("Command"))
                            {
                                if (null != subNewkey)
                                {
                                    subNewkey.SetValue(string.Empty, $@"{path} ""%1""");
                                    return true;
                                }
                                else
                                {
                                    ShowUnkownErrorMessage("null command key");
                                }
                            }
                        }
                        else
                        {
                            ShowUnkownErrorMessage("null file sub key");
                        }
                    }
                }
            }

            return false;
        }

        [SuppressMessage("ReSharper", "UnusedParameter.Local")]
        private static void ShowUnkownErrorMessage(string msg = "")
        {
            MessageBox.Show("An unkown error occured", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private static void ShowAdminRightsRequiredMessage()
        {
            MessageBox.Show("To install the application correctly, you must give admin rights to the application",
                "Warning",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
    }
}