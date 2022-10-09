using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dev.Framework.RevitUtils
{
    public class ExportUtils
    {
        #region revit内文件导出

        /// <summary>
        /// 图纸导出DWG
        /// </summary>
        /// <param name="doc"></param>
        public static void ExportToDWG(Document doc)
        {
            var viewSheets = new FilteredElementCollector(doc).OfClass(typeof(ViewSheet)).Cast<ViewSheet>().Where(vw =>
                           vw.ViewType == ViewType.DrawingSheet && !vw.IsTemplate);

            // create DWG export options
            DWGExportOptions dwgOptions = new DWGExportOptions
            {
                //MergedViews = true,
                //SharedCoords = true,
                FileVersion = ACADVersion.R2010
            };


            List<ElementId> views = new List<ElementId>();

            foreach (var sheet in viewSheets)
            {
                if (!sheet.IsPlaceholder)
                {
                    views.Add(sheet.Id);
                    if (views.Count > 10)
                    {
                        break;
                    }
                }
            }
            // For Web Deployment
            //doc.Export(@"D:\sheetExporterLocation", "TEST", views, dwgOptions);
            // For Local
            string resPath = Path.Combine(Directory.GetCurrentDirectory(), "ExportFile", "dwg");

            if (!Directory.Exists(resPath))
            {
                Directory.CreateDirectory(resPath);
            }
            doc.Export(resPath, "rvt", views, dwgOptions);
        }

        /// <summary>
        /// 图纸导出DWG
        /// </summary>
        /// <param name="doc"></param>
        public static void ExportToDWG(Document doc, List<int> ids, string path, string name = "rvt")
        {
            if (ids.Count == 0)
                return;
            List<ElementId> views = new List<ElementId>();
            ids.ForEach(e => views.Add(new ElementId(e)));

            // create DWG export options

            DWGExportOptions dwgOptions = DWGExportOptions.GetPredefinedOptions(doc, "A-全黑");
            dwgOptions.MergedViews = true;
            //dwgOptions.SharedCoords = true;
            dwgOptions.FileVersion = ACADVersion.R2010;
            //DWGExportOptions dwgOptions = new DWGExportOptions()
            //{
            //    MergedViews = true,
            //    //SharedCoords = true,
            //    FileVersion = ACADVersion.Default
            //};

            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            //name += "-" + (doc.GetElement(views.First()) as View).Name;
            doc.Export(path, name, views, dwgOptions);
            doc.Regenerate();
        }

        /// <summary>
        /// 官方api
        /// </summary>
        /// <param name="document"></param>
        /// <param name="view"></param>
        /// <param name="setupName"></param>
        /// <returns></returns>
        public static bool ExportToDWG(Document document, View view, string setupName)
        {
            bool exported = false;
            // Get the predefined setups and use the one with the given name.
            IList<string> setupNames = BaseExportOptions.GetPredefinedSetupNames(document);
            foreach (string name in setupNames)
            {
                if (name.CompareTo(setupName) == 0)
                {
                    // Export using the predefined options
                    DWGExportOptions dwgOptions = DWGExportOptions.GetPredefinedOptions(document, name);

                    // Export the active view
                    ICollection<ElementId> views = new List<ElementId>();
                    views.Add(view.Id);
                    // The document has to be saved already, therefore it has a valid PathName.
                    exported = document.Export(Path.GetDirectoryName(document.PathName),
                        Path.GetFileNameWithoutExtension(document.PathName), views, dwgOptions);
                    break;
                }
            }

            return exported;
        }

        /// <summary>
        /// 导出PDF文件
        /// </summary>
        /// <param name="datas"></param>
        public static void ExportPDF(Document doc, List<int> ids, string path)
        {

            ViewSet printableViews = new ViewSet();
            foreach (var id in ids)
            {
                View view = doc.GetElement(new ElementId(id)) as View;
                if (view != null && !view.IsTemplate && view.CanBePrinted)
                {
                    printableViews.Insert(view);
                }
            }

            //string url = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "test.pdf");
            //ProcessStartInfo info = new ProcessStartInfo(url);
            //info.Verb = "Print";
            //info.CreateNoWindow = true;


            PrintManager pm = doc.PrintManager;
            pm.PrintRange = PrintRange.Select;
            //ViewSheetSetting viewSheetSetting = pm.ViewSheetSetting;
            //viewSheetSetting.CurrentViewSheetSet.Views = printableViews;
            //viewSheetSetting.SaveAs("MyViewSet");
            //pm.CombinedFile = true;
            //pm.SubmitPrint();
            pm.SelectNewPrintDriver("Microsoft Print to PDF");
            pm.PrintToFile = true;
            pm.PrintToFileName = path;
            pm.CombinedFile = true;
            ///打印范围：所选视图/图纸
            //pm.ViewSheetSetting.InSession.Views = printableViews;
            //var currentPrintSetup = pm.PrintSetup.CurrentPrintSetting;
            //pm.PrintSetup.CurrentPrintSetting = pm.PrintSetup.InSession;
            //pm.PrintSetup.Save();
            pm.Apply();

            doc.Print(printableViews, true);
            //doc.Regenerate();
        }


        /// <summary>
        /// Revit文件另存为
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="path"></param>
        /// <returns></returns>

        public static bool RevitProjectSaveAs(Document doc, string path, string title)
        {
            bool isSuccessed = false;
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            SaveAsOptions options = new SaveAsOptions()
            {
                OverwriteExistingFile = true
            };

            string projectPath = Path.Combine(path, title + ".rvt");
            WorksharingSaveAsOptions worksharingSaveAsOptions = new WorksharingSaveAsOptions();
            try
            {
                //string resPath = Path.Combine(Directory.GetCurrentDirectory(), "ExportFile", "rvt");
                worksharingSaveAsOptions.SaveAsCentral = true;//保存为中心文件
                options.SetWorksharingOptions(worksharingSaveAsOptions);
                //options.PreviewViewId = doc.ActiveView.Id;
                doc.SaveAs(projectPath, options);
                isSuccessed = true;
            }
            catch (Exception ex)
            {
                worksharingSaveAsOptions.SaveAsCentral = false;
                options.SetWorksharingOptions(worksharingSaveAsOptions);
                doc.SaveAs(projectPath, options);
                isSuccessed = true;
            }
            return isSuccessed;
        }

        public static bool RevitProjectSaveAs(Document doc, string path, string title, bool isCenter = false)
        {
            bool isSuccessed = false;

            //string resPath = Path.Combine(Directory.GetCurrentDirectory(), "ExportFile", "rvt");
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            SaveAsOptions options = new SaveAsOptions()
            {
                OverwriteExistingFile = true
            };
            if (isCenter)
            {
                WorksharingSaveAsOptions worksharingSaveAsOptions = new WorksharingSaveAsOptions();
                worksharingSaveAsOptions.SaveAsCentral = true;//保存为中心文件
                options.SetWorksharingOptions(worksharingSaveAsOptions);
            }
            //options.PreviewViewId = doc.ActiveView.Id;

            string projectPath = Path.Combine(path, title + ".rvt");
            doc.SaveAs(projectPath, options);
            isSuccessed = true;

            return isSuccessed;
        }
        #endregion
    }
}
