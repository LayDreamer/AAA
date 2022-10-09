using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dev.Framework.RevitUtils
{
    public class FamilyHelper
    {
        #region 载入族

        /// <summary>
        /// 加载指定名称的族（默认地址）
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="familyname"></param>
        /// <param name="message"></param>
        public static void LoadFamilyFromLocal(Document doc,string path, List<string> familyNames, ref string message)
        {
            message += "族批量加载中..." + "\r\n";
            //var families = new List<Family>();
            //var files = Directory.GetFiles(FamilyPath, "*.rfa").ToList();
            var files = Directory.GetFiles(path, "*.rfa", SearchOption.AllDirectories).ToList();
            if (files.Count() == 0)
            {
                message += $"族文件夹为空" + "\r\n";
            }
            try
            {
                List<string> targetFamilyPaths = new List<string>();
                familyNames.ForEach(x =>
                {
                    var resFiles = files.FindAll(e => Path.GetFileName(e).Contains(x));
                    if (resFiles.Count() != 0)
                    {
                        targetFamilyPaths.AddRange(resFiles);
                    }
                });

                foreach (var familyPath in targetFamilyPaths)
                {
                    Family family = null;
                    if (File.Exists(familyPath))
                    {
                        bool success = doc.LoadFamily(familyPath, new FamilyLoadOptions(), out family);
                        string familyName = Path.GetFileNameWithoutExtension(familyPath);
                        if (!success)
                        {
                            message += $"项目中已经存在{familyName}" + "\r\n";
                        }
                        if (success)
                        {
                            message += $"项目中{familyName}载入完成" + "\r\n";
                        }
                    }
                }

                //foreach (var familyPath in files)
                //{
                //    Family family = null;
                //    if (File.Exists(familyPath))
                //    {
                //        if (familyNames.All(e => !familyPath.Contains(e)))
                //            continue;

                //        bool success = doc.LoadFamily(familyPath, new FamilyLoadOptions(), out family);
                //        string familyName = Path.GetFileNameWithoutExtension(familyPath);
                //        if (!success)
                //        {
                //            message += $"项目中已经存在{familyName}" + "\r\n";
                //        }
                //        if (success)
                //        {
                //            message += $"项目中{familyName}载入完成" + "\r\n";
                //        }
                //    }
                //}
            }
            catch (Exception ex)
            {
                message += $"族加载错误:{ex.Message}" + "\r\n";
            }
        }

        /// <summary>
        /// 加载文件夹中的族
        /// </summary>
        public static void LoadFamilyFromLocal(Document doc, string familyDirectory, ref string message)
        {
            message += "族批量加载中..." + "\r\n";
            var familyPathFiles = Directory.GetFiles(familyDirectory, "*.rfa", SearchOption.AllDirectories).ToList();
            if (familyPathFiles.Count() == 0)
            {
                message += $"族文件夹为空" + "\r\n";
            }
            try
            {
                foreach (var familyFilePath in familyPathFiles)
                {
                    Family family = null;
                    if (File.Exists(familyFilePath))
                    {
                        bool success = doc.LoadFamily(familyFilePath, new FamilyLoadOptions(), out family);
                        string familyName = Path.GetFileNameWithoutExtension(familyFilePath);
                        if (!success)
                        {
                            message += $"项目中已经存在{familyName}" + "\r\n";
                        }
                        if (success)
                        {
                            message += $"项目中{familyName}载入完成" + "\r\n";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                message += $"族加载错误:{ex.Message}" + "\r\n";
            }
        }

        /// <summary>
        /// 根据族文件地址载入
        /// </summary>
        public static void LoadFamilyByFilePath(Document doc, string filePath, ref string message)
        {
            //message += "族加载中..." + "\r\n";
            try
            {
                Family family = null;
                if (File.Exists(filePath))
                {
                    bool success = doc.LoadFamily(filePath, new FamilyLoadOptions(), out family);
                    string familyName = Path.GetFileNameWithoutExtension(filePath);
                    if (!success)
                    {
                        message += $"项目中已经存在族：【{familyName}】";
                    }
                    if (success)
                    {
                        message += $"项目中族：【{familyName}】载入完成";
                    }
                }
            }
            catch (Exception ex)
            {
                message += $"族加载错误:{ex.Message}";
            }
        }

        /// <summary>
        /// 加载本地族并返回族类型
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="familyname"></param>
        /// <param name="symbolName"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static FamilySymbol LoadFamilyFromLocal(Document doc,string path, string familyname, string symbolName = "")
        {
            FamilySymbol symbol = null;
            if (string.IsNullOrEmpty(symbolName))
                symbol = FindFamilySymbol(doc, familyname);
            else
                symbol = FindFamilySymbol(doc, familyname, symbolName);
            if (symbol != null)
                return symbol;
            string familyPath = Path.Combine(path, familyname + ".rfa");
            try
            {
                Family fs = null;
                if (File.Exists(familyPath))
                {
                    doc.LoadFamily(familyPath, new FamilyLoadOptions(), out fs);
                }
                if (fs != null)
                {
                    foreach (ElementId symbolId in fs.GetFamilySymbolIds())
                    {
                        Element elem = doc.GetElement(symbolId);
                        if (null != elem && (string.IsNullOrEmpty(symbolName) || elem.Name == symbolName))
                        {
                            symbol = elem as FamilySymbol;
                            break;
                        }
                    }
                    return symbol;
                }
                if (symbol == null)
                {
                    throw new Exception("族：【" + familyname + "】，类型：【" + symbolName + "】未找到");
                }
                return symbol;
            }
            catch (Exception ex)
            {
                throw new Exception("族：【" + familyname + "】，类型：【" + symbolName + "】未找到");
            }
            finally
            {
                if (symbol != null && !symbol.IsActive)
                    symbol.Activate();
            }
        }

        /// <summary>
        /// 通过族名称，在项目中查找已有的族
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="familyName"></param>
        /// <returns></returns>
        public static FamilySymbol FindFamilySymbol(Document doc, string familyName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            FilteredElementIterator itr = collector.OfClass(typeof(Family)).GetElementIterator();
            itr.Reset();
            while (itr.MoveNext())
            {
                Element elem = itr.Current;
                if (elem.GetType() != typeof(Family))
                {
                    continue;
                }
                if (elem.Name == familyName)
                {
                    Family family = (Family)elem;
                    foreach (ElementId symbolId in family.GetFamilySymbolIds())
                    {
                        FamilySymbol symbol = (FamilySymbol)doc.GetElement(symbolId);
                        return symbol;
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// 通过族名、类型查找相应的Symbol
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="familyName"></param>
        /// <param name="symbolName"></param>
        /// <returns></returns>
        public static FamilySymbol FindFamilySymbol(Document doc, string familyName, string symbolName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            FilteredElementIterator itr = collector.OfClass(typeof(Family)).GetElementIterator();
            itr.Reset();
            while (itr.MoveNext())
            {
                Element elem = itr.Current;
                if (elem.GetType() != typeof(Family))
                {
                    continue;
                }
                if (!string.Equals(elem.Name, familyName, StringComparison.CurrentCultureIgnoreCase))
                {
                    continue;
                }
                Family family = (Family)elem;
                foreach (ElementId symbolId in family.GetFamilySymbolIds())
                {
                    FamilySymbol symbol = (FamilySymbol)doc.GetElement(symbolId);
                    if (string.Equals(symbol.Name, symbolName, StringComparison.CurrentCultureIgnoreCase))
                    {
                        return symbol;
                    }
                }
            }
            return null;
        }
        #endregion
    }

    /// <summary>
    /// 载入族设置，覆盖构建族参数
    /// </summary>
    public class FamilyLoadOptions : IFamilyLoadOptions
    {
        public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
        {
            overwriteParameterValues = true;
            return true;
        }

        public bool OnSharedFamilyFound(Family sharedFamily, bool familyInUse, out FamilySource source, out bool overwriteParameterValues)
        {
            source = FamilySource.Family;
            overwriteParameterValues = true;
            return true;
        }
    }
}
