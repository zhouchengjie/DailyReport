using System;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Web;

using System.Globalization;
using System.Collections.Generic;
using System.Reflection;
using System.IO;
using System.Collections;
using System.Text;
using System.Text.RegularExpressions;
using System.Linq;
using System.Security.AccessControl;
using System.Net;

namespace DailyReport
{
    /// <summary>
    /// Summary description for CommonFunction
    /// </summary>
    public class CommonFunction
    {
        public CommonFunction()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        public static decimal GetDeciamlFromCurrency(string currency, bool dividePercentage)
        {
            decimal rtn = GetDeciamlFromCurrency(currency);
            if (currency.Contains("%") == true)
            {
                if (dividePercentage == true)
                {
                    rtn = rtn / 100;
                }
            }
            return rtn;
        }
        public static decimal GetDeciamlFromCurrency(string currency)
        {
            decimal amount = 0;
            string stramount = currency;
            stramount = stramount.Replace("$", "");
            stramount = stramount.Replace("%", "");
            stramount = stramount.Replace(",", "");
            decimal.TryParse(stramount, out amount);
            return amount;
        }

        public static string GetNewPathForDuplicates(string path)
        {
            string directory = Path.GetDirectoryName(path);
            string filename = Path.GetFileNameWithoutExtension(path);
            string extension = Path.GetExtension(path);
            int counter = 1;
            string newFullPath = path;
            while (System.IO.File.Exists(newFullPath))
            {
                string newFilename = string.Format("{0}({1}){2}", filename, counter, extension);
                newFullPath = Path.Combine(directory, newFilename);
                counter++;
            }
            return newFullPath;
        }

        public static string GetStringValue(object obj)
        {
            if (obj == null)
            {
                return null;
            }
            return obj.ToString();
        }

        public static void StopProcess(string processName)
        {
            try
            {
                System.Diagnostics.Process[] ps = System.Diagnostics.Process.GetProcessesByName(processName);
                foreach (System.Diagnostics.Process p in ps)
                {
                    p.Kill();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void AddSharePermissionToFolder(string folderPath, string domain, string accountName)
        {
            DirectoryInfo dir = new DirectoryInfo(folderPath);
            string fullUserName = domain + "\\" + accountName;
            //获得该文件夹的所有访问权限
            System.Security.AccessControl.DirectorySecurity dirSecurity = dir.GetAccessControl(AccessControlSections.All);
            InheritanceFlags inherits = InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit;
            //添加Users用户组的访问权限规则 完全控制权限
            //FileSystemAccessRule usersFileSystemAccessRule = new FileSystemAccessRule("Users", FileSystemRights.FullControl, inherits, PropagationFlags.None, AccessControlType.Allow);
            FileSystemAccessRule userRule = new FileSystemAccessRule(fullUserName, FileSystemRights.FullControl, inherits, PropagationFlags.None, AccessControlType.Allow);
            bool isModified = false;
            dirSecurity.ModifyAccessRule(AccessControlModification.Add, userRule, out isModified);
            //设置访问权限
            dir.SetAccessControl(dirSecurity);
        }

        public static int GetTimeToSeconds(string time) 
        {
            string [] timePart = time.Split(new char[] { ':' });
            if (timePart == null || timePart.Length != 3) 
            {
                return 0;
            }
            int hour = 0;
            int minute = 0;
            int second = 0;
            int.TryParse(timePart[0], out hour);
            int.TryParse(timePart[1], out minute);
            int.TryParse(timePart[2], out second);

            return 3600 * hour + 60 * minute + second;
        }

        public static string GetIPAddress() 
        {
            string addressIP = string.Empty;
            foreach (IPAddress _IPAddress in Dns.GetHostEntry(Dns.GetHostName()).AddressList)
            {
                if (_IPAddress.AddressFamily.ToString() == "InterNetwork")
                {
                    addressIP = _IPAddress.ToString();
                }
            }
            return addressIP;
        }
    }
}