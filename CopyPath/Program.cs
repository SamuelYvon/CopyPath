using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Security.Permissions;
using System.Windows.Forms;

namespace CopyPath
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            if (args.Length > 0)
            {
                Clipboard.SetText(args[0]);
            }
            else
            {
                string filePath = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName;
                AddLinkToFiles(filePath);
            }
        }

        const string SHELL_SUBKEY = @"*\shell";
        const string OPERATION_NAME = "Copy full path";

        [PrincipalPermission(SecurityAction.Assert, Role = @"BUILTIN\Administrators")]
        static void AddLinkToFiles(string path)
        {
            using (RegistryKey shellKey = Registry.ClassesRoot.OpenSubKey(SHELL_SUBKEY, true))
            {
                string[] shellKeyChildsName = shellKey.GetSubKeyNames();

                ISet<string> shellKeyChilds = new HashSet<string>(shellKeyChildsName);

                if (!shellKeyChilds.Contains(OPERATION_NAME))
                {
                    using (RegistryKey newkey = shellKey.CreateSubKey(OPERATION_NAME))
                    using (RegistryKey subNewkey = newkey.CreateSubKey("Command"))
                    {
                        subNewkey.SetValue("", string.Format(@"{0} ""%1""", path));
                    }
                }
            }
        }
    }
}
