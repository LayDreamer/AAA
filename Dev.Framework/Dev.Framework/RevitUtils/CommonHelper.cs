using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dev.Framework.RevitUtils
{
    public static class CommonHelper
    {
        #region 元素参数赋值

        #region get

        public static string GetStringValue(this Element instance, BuiltInParameter builtInParameter)
        {
            string value = string.Empty;
            if (instance != null)
            {
                Parameter para = instance.get_Parameter(builtInParameter);

                if (para != null && para.StorageType == StorageType.String)
                    using (para)
                        value = para.AsString();
            }

            return value;
        }
        public static string GetStringValue(this Element instance, string strName)
        {
            string value = string.Empty;

            if (instance != null && !string.IsNullOrEmpty(strName))
            {
                Parameter para = instance.LookupParameter(strName);
                if (para != null && para.StorageType == StorageType.String)
                    using (para)
                        value = para.AsString();
            }

            return value;
        }

        public static ElementId GetElementId(this Element instance, BuiltInParameter builtInParameter)
        {
            ElementId value = ElementId.InvalidElementId;
            if (instance != null)
            {
                Parameter para = instance.get_Parameter(builtInParameter);
                if (para != null && para.StorageType == StorageType.ElementId)
                    using (para)
                        value = para.AsElementId();
            }

            return value;
        }
        public static ElementId GetElementId(this Element instance, string strName)
        {
            ElementId value = ElementId.InvalidElementId;

            if (instance != null && !string.IsNullOrEmpty(strName))
            {
                Parameter para = instance.LookupParameter(strName);
                if (para != null && para.StorageType == StorageType.ElementId)
                    using (para)
                        value = para.AsElementId();
            }

            return value;
        }
        public static double GetDoubleValue(this Element instance, BuiltInParameter builtInParameter)
        {
            double value = 0.0;
            if (instance != null)
            {
                Parameter para = instance.get_Parameter(builtInParameter);
                if (para != null && para.StorageType == StorageType.Double)
                    using (para)
                        value = para.AsDouble();
            }

            return value;
        }
        public static string GetValueAsString(this Element instance, BuiltInParameter builtInParameter, FormatOptions formatOptions = null)
        {
            string value = string.Empty;

            if (instance != null)
            {
                Parameter para = instance.get_Parameter(builtInParameter);
                if (para != null)
                    using (para)
                        value = formatOptions == null ? para.AsValueString() : para.AsValueString(formatOptions);
            }

            return value;
        }
        public static string GetValueAsString(this Element instance, string strName, FormatOptions formatOptions = null)
        {
            string value = string.Empty;

            if (instance != null)
            {
                Parameter para = instance.LookupParameter(strName);
                if (para != null)
                    using (para)
                        value = formatOptions == null ? para.AsValueString() : para.AsValueString(formatOptions);
            }

            return value;
        }
        public static int GetInteger(this Element instance, BuiltInParameter builtInParameter)
        {
            int value = 0;
            if (instance != null)
            {
                Parameter para = instance.get_Parameter(builtInParameter);
                if (para != null && para.StorageType == StorageType.Integer)
                    using (para)
                        value = para.AsInteger();
            }

            return value;
        }
        public static int GetInteger(this Element instance, string strName)
        {
            int value = 0;
            if (instance != null && !string.IsNullOrEmpty(strName))
            {
                Parameter para = instance.LookupParameter(strName);
                if (para != null && para.StorageType == StorageType.Integer)
                    using (para)
                        value = para.AsInteger();
            }

            return value;
        }

        public static double GetDoubleValue(this Element instance, string strName)
        {
            double value = 0.0;
            if (instance != null && !string.IsNullOrEmpty(strName))
            {
                Parameter para = instance.LookupParameter(strName);
                if (para != null && para.StorageType == StorageType.Double)
                    using (para)
                        value = para.AsDouble();
            }

            return value;
        }

        #endregion

        #region set

        public static void SetDoubleValue(this Element instance, BuiltInParameter builtInParameter, double value)
        {
            if (instance != null)
            {
                var param = instance.get_Parameter(builtInParameter);
                if (
                    param != null &&
                    param.StorageType == StorageType.Double &&
                    !param.IsReadOnly)
                    using (param)
                        param.Set(value);
            }
        }
        public static void SetDoubleValue(this Element instance, string strName, double value)
        {
            if (instance != null && !string.IsNullOrEmpty(strName))
            {
                var param = instance.LookupParameter(strName);
                if (
                    param != null &&
                    param.StorageType == StorageType.Double &&
                    !param.IsReadOnly)
                    using (param)
                        param.Set(value);
            }
        }

        public static void SetInteger(this Element instance, string strName, int value)
        {
            if (instance != null && !string.IsNullOrEmpty(strName))
            {
                var param = instance.LookupParameter(strName);
                if (
                    param != null &&
                    param.StorageType == StorageType.Integer &&
                    !param.IsReadOnly)
                    using (param)
                        param.Set(value);
            }
        }
        public static void SetStringValue(this Element instance, string strName, string value)
        {
            if (instance != null && !string.IsNullOrEmpty(strName))
            {
                var param = instance.LookupParameter(strName);
                if (
                    param != null &&
                    param.StorageType == StorageType.String &&
                    !param.IsReadOnly)
                    using (param)
                        param.Set(value);
            }
        }
        public static void SetStringValue(this Element instance, BuiltInParameter builtInParameter, string value)
        {
            if (instance != null)
            {
                var param = instance.get_Parameter(builtInParameter);
                if (
                    param != null &&
                    param.StorageType == StorageType.String &&
                    !param.IsReadOnly)
                    using (param)
                        param.Set(value);
            }
        }
        public static void SetInteger(this Element instance, BuiltInParameter builtInParameter, int value)
        {
            if (instance != null)
            {
                var param = instance.get_Parameter(builtInParameter);
                if (
                    param != null &&
                    param.StorageType == StorageType.Integer &&
                    !param.IsReadOnly)
                    using (param)
                        param.Set(value);
            }
        }
        public static void SetElmentId(this Element instance, BuiltInParameter builtInParameter, ElementId value)
        {
            if (instance != null)
            {
                var param = instance.get_Parameter(builtInParameter);
                if (
                    param != null &&
                    param.StorageType == StorageType.ElementId &&
                    !param.IsReadOnly)
                    using (param)
                        param.Set(value);
            }
        }
        public static void SetElmentId(this Element instance, string strName, ElementId value)
        {
            if (instance != null && !string.IsNullOrEmpty(strName))
            {
                var param = instance.LookupParameter(strName);
                if (
                    param != null &&
                    param.StorageType == StorageType.ElementId &&
                    !param.IsReadOnly)
                    using (param)
                        param.Set(value);
            }
        }


        #endregion

        #endregion

        #region Revit移除提醒框

        public class FailureHandleForFunction : IFailuresPreprocessor
        {
            public string ErrorMessage { set; get; }
            public string ErrorSeverity { set; get; }
            public FailureHandleForFunction()
            {
                ErrorMessage = "";
                ErrorSeverity = "";
            }
            public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
            {
                Document doc = failuresAccessor.GetDocument();

                IList<FailureMessageAccessor> failureMessages = failuresAccessor.GetFailureMessages();
                foreach (FailureMessageAccessor failureMessageAccessor in failureMessages)
                {
                    try
                    {
                        FailureSeverity failureSeverity = failureMessageAccessor.GetSeverity();
                        ErrorSeverity = failureSeverity.ToString();
                        if (failureSeverity == FailureSeverity.Warning)
                        {
                            failureMessageAccessor.GetDefaultResolutionCaption();
                            // 如果是警告，则禁止消息框  
                            failuresAccessor.DeleteWarning(failureMessageAccessor);
                        }
                        else
                        {
                            ErrorMessage += failureMessageAccessor.GetDescriptionText();
                            foreach (ElementId eid in failureMessageAccessor.GetFailingElementIds())
                            {
                                ErrorMessage += doc.GetElement(eid).Name + "ID:" + eid.ToString();
                            }
                            if (failureMessageAccessor.HasResolutions())
                                failuresAccessor.ResolveFailure(failureMessageAccessor);
                            return FailureProcessingResult.ProceedWithCommit;
                        }
                    }
                    catch
                    {
                    }
                }
                return FailureProcessingResult.Continue;
            }
            public static FailureHandleForFunction SetFailureHandle(Transaction transaction)
            {
                FailureHandlingOptions failureHandlingOptions = transaction.GetFailureHandlingOptions();

                FailureHandleForFunction failureHandler = new FailureHandleForFunction();

                failureHandlingOptions.SetFailuresPreprocessor(failureHandler);
                failureHandlingOptions.SetClearAfterRollback(true);

                transaction.SetFailureHandlingOptions(failureHandlingOptions);
                return failureHandler;
            }
        }

        public class RemoveWarmingPreprocessor : IFailuresPreprocessor
        {
            RemoveWarmingPreprocessor() { }
            public static readonly IFailuresPreprocessor Instance = new RemoveWarmingPreprocessor();
            public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
            {
                var fmas = failuresAccessor.GetFailureMessages();
                if (fmas.Count > 0)
                {
                    var transactionName = failuresAccessor.GetTransactionName();
                    failuresAccessor.DeleteAllWarnings();

                    return FailureProcessingResult.ProceedWithCommit;
                }

                return FailureProcessingResult.Continue;
            }
        }

        public class WarmingPreprocessor : IFailuresPreprocessor
        {
            public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
            {
                IList<FailureMessageAccessor> fmas = failuresAccessor.GetFailureMessages();
                bool bContinue = false;
                if (fmas.Count > 0)
                {
                    foreach (FailureMessageAccessor fma in fmas)
                    {
                        string strMsg = fma.GetDescriptionText().ToString();
                        if (strMsg == "无法使图元保持连接。")
                        {
                            bContinue = true;
                            failuresAccessor.ResolveFailure(fma);
                        }
                        else if (strMsg.Contains("线太短"))
                        {
                            var ids = fma.GetFailingElementIds().ToList();
                            failuresAccessor.DeleteElements(ids);
                        }
                        else
                        {
                            //删除警告
                            failuresAccessor.DeleteWarning(fma);
                        }
                    }

                }
                if (bContinue)
                {
                    return FailureProcessingResult.ProceedWithCommit;
                }
                else
                {
                    return FailureProcessingResult.Continue;
                }
            }
        }

        public class RevitFailureHandler : IFailuresPreprocessor
        {
            public string ErrorMessage { set; get; }
            public string ErrorSeverity { set; get; }

            public RevitFailureHandler()
            {
                ErrorMessage = "";
                ErrorSeverity = "";
            }

            public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
            {
                IList<FailureMessageAccessor> failureMessages = failuresAccessor.GetFailureMessages();
                foreach (FailureMessageAccessor failureMessageAccessor in failureMessages)
                {
                    FailureDefinitionId id = failureMessageAccessor.GetFailureDefinitionId();
                    try
                    {
                        ErrorMessage = failureMessageAccessor.GetDescriptionText();
                    }
                    catch
                    {
                        ErrorMessage = "Unknown Error";
                    }
                    try
                    {
                        FailureSeverity failureSeverity = failureMessageAccessor.GetSeverity();
                        ErrorSeverity = failureSeverity.ToString();
                        if (failureSeverity == FailureSeverity.Warning)
                        {
                            //if (failureMessageAccessor.GetDescriptionText() == "同一位置处具有相同实例。这将导致在明细表中重复计算。.")
                            //{
                            //    List<ElementId> c = failureMessageAccessor.GetFailingElementIds().ToList();
                            //    if (c.Count > 1)
                            //    {
                            //        failuresAccessor.DeleteElements(new List<ElementId>() { c[1] });
                            //    }
                            //}
                            //else
                            //{
                            // 如果是警告，则禁止消息框  
                            failureMessageAccessor.GetDefaultResolutionCaption();
                            //}
                            failuresAccessor.DeleteWarning(failureMessageAccessor);
                        }
                        else
                        {
                            // 如果是错误：则取消导致错误的操作，但是仍然继续整个事务  
                            if (ErrorMessage.Contains("线太短"))
                            {
                                failureMessageAccessor.SetCurrentResolutionType(FailureResolutionType.DeleteElements);
                                failuresAccessor.ResolveFailure(failureMessageAccessor);
                                return FailureProcessingResult.ProceedWithCommit;
                            }
                            //else
                            //{
                            //    failuresAccessor.DeleteWarning(failureMessageAccessor);
                            //    return FailureProcessingResult.ProceedWithCommit;
                            //}
                        }
                    }
                    catch (Exception ex)
                    {
                    }
                }
                return FailureProcessingResult.Continue;
            }

            [Description("这个方法用在事务开始前，在FailureHandler初始化后调用")]
            /// <summary>
            /// 这个方法用在事务开始前，在FailureHandler初始化后调用
            /// </summary>
            public static void SetFailedHandlerBeforeTransaction(IFailuresPreprocessor failureHandler, Transaction transaction)
            {
                FailureHandlingOptions failureHandlingOptions = transaction.GetFailureHandlingOptions();
                failureHandlingOptions.SetFailuresPreprocessor(failureHandler);
                // 这句话是关键  
                //failureHandlingOptions.SetClearAfterRollback(true);
                transaction.SetFailureHandlingOptions(failureHandlingOptions);
            }
        }
        #endregion

    }
}
