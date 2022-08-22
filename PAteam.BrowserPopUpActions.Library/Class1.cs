using Direct.Shared;
using System;
using log4net;
using System.Windows.Automation;
using System.Text;

namespace Direct.BrowserDialogActions.Library
{
    [DirectSealed]
    [DirectDom("Browser Dialog Actions")]
    [ParameterType(false)]
    public static class BrowserDialogFunctions
    {
        private static readonly ILog _log = LogManager.GetLogger("LibraryObjects");
        private static string packageName = "Direct.BrowserDialogActions.Library";
        private static PropertyCondition DialogPropertyCondition = new PropertyCondition(AutomationElement.LocalizedControlTypeProperty, "dialog");
        private static PropertyCondition inputFileButtonCondition = new PropertyCondition(AutomationElement.AutomationIdProperty, "file-upload-button");
        private static PropertyCondition DialogOKButtonCondition = new PropertyCondition(AutomationElement.NameProperty, "OK");
        private static PropertyCondition DialogCancelButtonCondition = new PropertyCondition(AutomationElement.NameProperty, "Cancel");

        private static void LogDebug(string message)
        {
            if (_log.IsDebugEnabled)
            {
                _log.Debug(packageName + " - " + message);
            }
        }

        private static void LogError(string message, Exception exception)
        {
            if (_log.IsErrorEnabled)
            {
                _log.Debug(packageName + " - " + message, exception);
            }
        }

        [DirectDom("Get Dialog Existence")]
        [DirectDomMethod("Get Active Tab Dialog Existence")]
        [MethodDescription("Evaluates if the browser Dialog windows like: alert, confirm could be found")]
        public static bool DialogExists()
        {
            try
            {
                return GetDialogExistence();
            }
            catch (Exception e)
            {
                LogError("Get Browser Dialog Existence Exception", e);
                return false;
            }
        }

        [DirectDom("Get Dialog Text")]
        [DirectDomMethod("Get Active Tab Dialog Text")]
        [MethodDescription("Returns browser Dialog text on active tab if found")]
        public static string GetDialogText()
        {
            try
            {
                AutomationElement DialogAUElement = GetBrowserChildAUElement(DialogPropertyCondition, TreeScope.Children);
                if (DialogAUElement != null)
                {
                    Condition LabelPropertyCondition = new PropertyCondition(AutomationElement.LocalizedControlTypeProperty, "text");
                    LogDebug("Finding browser Dialog text elements...");
                    AutomationElementCollection TextElements = DialogAUElement.FindAll(TreeScope.Descendants, LabelPropertyCondition);
                    StringBuilder finalText = new StringBuilder();
                    LogDebug("Found " + TextElements.Count.ToString() + " text elements");
                    foreach (AutomationElement TextElement in TextElements)
                    {
                        LogDebug("Retrieving text property from found text element");
                        finalText.AppendLine(GetAUElementText(TextElement));
                    }

                    return finalText.ToString();
                }

                return string.Empty;
            }
            catch (Exception e)
            {
                LogError("Get Browser Dialog Text Exception", e);
                return string.Empty;
            }
        }

        [DirectDom("Click Dialog OK Button")]
        [DirectDomMethod("Click Active Tab Dialog OK Button")]
        [MethodDescription("Clicks OK Button on found browser active tab pop up")]
        public static bool ClickDialogOK()
        {
            return ClickDialogButton(DialogOKButtonCondition);
        }

        [DirectDom("Click Dialog Cancel Button")]
        [DirectDomMethod("Click Active Tab Dialog Cancel Button")]
        [MethodDescription("Clicks Cancel Button on found browser active tab pop up")]
        public static bool ClickDialogCancel()
        {
            return ClickDialogButton(DialogCancelButtonCondition);
        }



        [DirectDom("Open File Dialog on active page")]
        [DirectDomMethod("Open File Dialog On Active Page")]
        [MethodDescription("Opens File Dialog to select files to upload")]
        public static bool OpenFileDialog()
        {
            try
            {
                AutomationElement InputFileAUElement = GetBrowserChildAUElement(inputFileButtonCondition, TreeScope.Descendants);
                if (InputFileAUElement != null)
                {
                    LogDebug("Clicking file dialog open button...");
                    ClickOnButton(InputFileAUElement);
                    return true;
                }

                return false;
            }
            catch (Exception e)
            {
                LogError("OpenFileDialog", e);
                return false;
            }
        }

        private static bool GetDialogExistence()
        {
            LogDebug("Finding browser Dialog element...");
            AutomationElement DialogAUElement = GetBrowserChildAUElement(DialogPropertyCondition, TreeScope.Children);

            if (DialogAUElement != null)
            {
                LogDebug("Found dialog with localized control type: " + DialogAUElement.Current.LocalizedControlType + " and name: " + DialogAUElement.Current.Name);
                return true;
            }

            return false;
        }

        private static AutomationElement GetBrowserChildAUElement(PropertyCondition elementCondition, TreeScope scope)
        {
            AutomationElementCollection BrowserAUElements = GetBrowserAUElements();
            if (BrowserAUElements.Count == 0)
            {
                LogDebug("Browser elements not found! Please check if browser is open and active");
                return null;
            }

            LogDebug("BrowserAUElements Count: " + BrowserAUElements.Count.ToString());
            Condition BrowserRootPropertyCondition = new PropertyCondition(AutomationElement.ClassNameProperty, "BrowserRootView");
            foreach (AutomationElement BrowserAUElement in BrowserAUElements)
            {
                AutomationElement BrowserRootAUElement = BrowserAUElement.FindFirst(TreeScope.Children, BrowserRootPropertyCondition);
                if (BrowserRootAUElement != null)
                {
                    AutomationElement AUElement = BrowserRootAUElement.FindFirst(scope, elementCondition);
                    if (AUElement != null)
                    {
                        return AUElement;
                    }
                }
            }
            return null;
        }



        private static AutomationElementCollection GetBrowserAUElements()
        {
            Condition BrowserPropertyCondition = new PropertyCondition(AutomationElement.ClassNameProperty, "Chrome_WidgetWin_1");
            AutomationElement RootElement = GetRootAUElement();
            if (RootElement == null)
            {
                LogDebug("Root element not found!");
                return null;

            }
            LogDebug("Finding all browser elements...");
            return RootElement.FindAll(TreeScope.Children, BrowserPropertyCondition);
        }

        private static AutomationElement GetRootAUElement()
        {
            LogDebug("Finding Root Element...");
            return AutomationElement.RootElement;
        }

        private static void ClickOnButton(AutomationElement button)
        {
            object invokePattern = null;
            button.TryGetCurrentPattern(InvokePattern.Pattern, out invokePattern);
            ((InvokePattern)invokePattern).Invoke();
        }

        private static bool ClickDialogButton(PropertyCondition DialogButtonCondition)
        {
            try
            {
                AutomationElement DialogAUElement = GetBrowserChildAUElement(DialogPropertyCondition, TreeScope.Children);
                if (DialogAUElement != null)
                {
                    LogDebug("Finding browser Dialog button element...");
                    AutomationElement ButtonElement = DialogAUElement.FindFirst(TreeScope.Descendants, DialogButtonCondition);
                    LogDebug("Clicking OK Button...");
                    ClickOnButton(ButtonElement);
                    return true;
                }

                return false;
            }
            catch (Exception e)
            {
                LogError("Close Browser Dialog", e);
                return false;
            }
        }

        public static string GetAUElementText(AutomationElement element)
        {
            object patternObj;
            if (element.TryGetCurrentPattern(ValuePattern.Pattern, out patternObj))
            {
                var valuePattern = (ValuePattern)patternObj;
                return valuePattern.Current.Value;
            }
            else if (element.TryGetCurrentPattern(TextPattern.Pattern, out patternObj))
            {
                var textPattern = (TextPattern)patternObj;
                return textPattern.DocumentRange.GetText(-1).TrimEnd('\r');
            }
            else
            {
                return element.Current.Name;
            }
        }
    }
}