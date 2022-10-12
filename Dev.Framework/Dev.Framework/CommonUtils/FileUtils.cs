using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dev.Framework.CommonUtils
{
    public class FileUtils
    {
        //桌面路径
        public static string DesktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);


        #region 打开

        /// <summary>
        /// 打开文件
        /// </summary>
        /// <param name="filePathAndName">文件的路径和名称（比如：C:\Users\Administrator\test.txt）</param>
        /// <param name="isWaitFileClose">是否等待文件关闭（true：表示等待）</param>
        public static void OpenFile(string filePathAndName, bool isWaitFileClose = false)
        {
            Process process = new Process();
            ProcessStartInfo psi = new ProcessStartInfo(filePathAndName);
            process.StartInfo = psi;

            process.StartInfo.UseShellExecute = true;

            try
            {
                process.Start();

                //等待打开的程序关闭
                if (isWaitFileClose)
                {
                    process.WaitForExit();
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                process?.Close();

            }
        }

        /// <summary>
        /// 打开目录并选中文件
        /// </summary>
        /// <param name="filePathAndName"></param>
        public static void OpenFloderAndSelectedFile(string filePathAndName)
        {
            if (string.IsNullOrEmpty(filePathAndName) || !File.Exists(filePathAndName))
                return;
            Process process = new Process();
            ProcessStartInfo psi = new ProcessStartInfo("Explorer.exe");
            psi.Arguments = "/e,/select," + filePathAndName;
            process.StartInfo = psi;
            try
            {
                process.Start();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                process?.Close();
            }
        }

        #endregion


        /// <summary>
        /// 获取文件中的文本
        /// </summary>
        /// <param name="sPath"></param>
        /// <returns></returns>
        public static string GetTextFromFile(string sPath)
        {
            string InvalidPathChar = MakeValidPathName(sPath);
            if (!InvalidPathChar.Equals("")) return null;

            List<string> tmp = new List<string>();
            string Allline;
            StreamReader sr = new StreamReader(sPath, Encoding.UTF8);
            Allline = sr.ReadToEnd();
            sr.Close();
            return Allline;
        }


        /// <summary>
        /// 文件夹路径中是否包含不能包含的字符
        /// </summary>
        /// <param name="spath">文件夹路径</param>
        /// <returns>不能包含的字符</returns>
        public static string MakeValidPathName(string spath)
        {
            var invalidPathNameChars = Path.GetInvalidPathChars();
            foreach (var c in spath)
            {
                if (invalidPathNameChars.Contains(c))
                {
                    return c.ToString();
                }
            }
            return "";
        }


        /// <summary>
        /// 重名文件自动添加新文件名
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static string GetNewPathForDupes(string path)
        {
            string newFullPath = path.Trim();
            if (System.IO.File.Exists(path))
            {
                string directory = Path.GetDirectoryName(path);
                string filename = Path.GetFileNameWithoutExtension(path);
                string extension = Path.GetExtension(path);
                int counter = 1;
                do
                {
                    string newFilename = string.Format("{0}({1}){2}", filename, counter, extension);
                    newFullPath = Path.Combine(directory, newFilename);
                    counter++;
                } while (System.IO.File.Exists(newFullPath));
            }
            return newFullPath;
        }

        /// <summary>
        /// 复制文件
        /// </summary>
        /// <param name="srcPath">源地址</param>
        /// <param name="destPath">目标地址</param>
        public static void CopyDirectory(string srcPath, string destPath)
        {
            DirectoryInfo dir = new DirectoryInfo(srcPath);
            FileSystemInfo[] fileinfo = dir.GetFileSystemInfos();
            foreach (FileSystemInfo i in fileinfo)
            {
                if (i is DirectoryInfo)
                {     //判断是否文件夹
                    if (!Directory.Exists(destPath + "\\" + i.Name))
                    {
                        Directory.CreateDirectory(destPath + "\\" + i.Name);
                    }
                    CopyDirectory(i.FullName, destPath + "\\" + i.Name);
                }
                else
                {
                    File.Copy(i.FullName, destPath + "\\" + i.Name, true);
                }
            }
        }


        /// <summary>
        /// 删除指定目录下所有文件
        /// </summary>
        /// <param name="srcPath"></param>
        public static void DeleteDirectory(string target_dir)
        {
            string[] files = Directory.GetFiles(target_dir);
            string[] dirs = Directory.GetDirectories(target_dir);
            foreach (string file in files)
            {
                File.SetAttributes(file, FileAttributes.Normal);
                File.Delete(file);
            }
            foreach (string dir in dirs)
            {
                DeleteDirectory(dir);
            }
            Directory.Delete(target_dir, false);
        }



        #region 共享文件夹操作

        /// <summary>
        /// 创建共享文件夹远程连接
        /// </summary>
        public static bool ConnectState(string path, string userName, string passWord)
        {
            bool Flag = false;
            Process proc = new Process();
            try
            {
                proc.StartInfo.FileName = "cmd.exe";
                proc.StartInfo.UseShellExecute = false;
                proc.StartInfo.RedirectStandardInput = true;
                proc.StartInfo.RedirectStandardOutput = true;
                proc.StartInfo.RedirectStandardError = true;
                proc.StartInfo.CreateNoWindow = true;
                proc.Start();
                string dosLine = "start" + path + " ";
                proc.StandardInput.WriteLine(dosLine);
                proc.StandardInput.WriteLine("exit");
                while (!proc.HasExited)
                {
                    proc.WaitForExit(1000);
                }
                string errormsg = proc.StandardError.ReadToEnd();
                proc.StandardError.Close();
                if (string.IsNullOrEmpty(errormsg))
                {
                    Flag = true;
                }
                else
                {
                    throw new Exception(errormsg);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                proc.Close();
                proc.Dispose();
            }
            return Flag;
        }

        #endregion

        /// <summary>
        /// 写入本地文本文件
        /// </summary>
        /// <param name="path">地址</param>
        /// <param name="data">数据</param>
        public static void WriteTextToLocal(string path, string data)
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
            using (StreamWriter output = File.CreateText(path))
            {
                output.WriteLine(data);
                output.Close();
            }
        }
    }
}
