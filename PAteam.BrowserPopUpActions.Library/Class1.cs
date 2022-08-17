using Direct.Shared;
using System;
using log4net;
using System.Windows.Automation;
using System.Text;

namespace Direct.BrowserPopUpActions.Library
{
    [DirectSealed]
    [DirectDom("Browser PopUp Functions")]
    [ParameterType(false)]
    public static class BrowserPopUpFunctions
    {
        private static readonly ILog _log = LogManager.GetLogger("LibraryObjects");
        private static string packageName = "Direct.BrowserPopUpActions.Library";
        private static PropertyCondition popUpPropertyCondition = new PropertyCondition(AutomationElement.LocalizedControlTypeProperty, "dialog");
        private static PropertyCondition inputFileButtonCondition = new PropertyCondition(AutomationElement.AutomationIdProperty, "file-upload-button");
        private static PropertyCondition popupOKButtonCondition = new PropertyCondition(AutomationElement.NameProperty, "OK");
        private static PropertyCondition popupCancelButtonCondition = new PropertyCondition(AutomationElement.NameProperty, "Cancel");

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

        [DirectDom("Get Pop-up Existence")]
        [DirectDomMethod("Get Browser Pop-up Existence")]
        [MethodDescription("Iterates over all browser windows and evaluate if the popup notification could be found")]
        public static bool PopUpExists()
        {
            try
            {
                return GetPopUpExistence();
            }
            catch (Exception e)
            {
                LogError("Get Browser PopUp Existence Exception", e);
                return false;
            }
        }

        [DirectDom("Get Pop-up Text")]
        [DirectDomMethod("Get Browser Pop-up Text")]
        [MethodDescription("Iterates over all browser windows and return text from popup window if found")]
        public static string GetPopUpText()
        {
            try
            {
                AutomationElement PopUpAUElement = GetBrowserChildAUElement(popUpPropertyCondition, TreeScope.Children);
                if (PopUpAUElement != null)
                {
                    Condition LabelPropertyCondition = new PropertyCondition(AutomationElement.LocalizedControlTypeProperty, "text");
                    LogDebug("Finding browser popup text elements...");
                    AutomationElementCollection TextElements = PopUpAUElement.FindAll(TreeScope.Descendants, LabelPropertyCondition);
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
                LogError("Get Browser PopUp Text Exception", e);
                return string.Empty;
            }
        }

        [DirectDom("Click Pop-up OK Button")]
        [DirectDomMethod("Click Pop-up OK Button")]
        [MethodDescription("Clicks OK Button on found browser active tab pop up")]
        public static bool ClickPopUpOK()
        {
            return ClickPopUpButton(popupOKButtonCondition);
        }

        [DirectDom("Click Pop-up Cancel Button")]
        [DirectDomMethod("Click Pop-up Cancel Button")]
        [MethodDescription("Clicks Cancel Button on found browser active tab pop up")]
        public static bool ClickPopUpCancel()
        {
            return ClickPopUpButton(popupCancelButtonCondition);
        }



        [DirectDom("Open File Dialog on active page")]
        [DirectDomMethod("Open file dialog on active page")]
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

        private static bool GetPopUpExistence()
        {
            LogDebug("Finding browser popup element...");
            AutomationElement PopUpAUElement = GetBrowserChildAUElement(popUpPropertyCondition, TreeScope.Children);

            if (PopUpAUElement != null)
            {
                LogDebug("Found dialog with localized control type: " + PopUpAUElement.Current.LocalizedControlType + " and name: " + PopUpAUElement.Current.Name);
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

        private static bool ClickPopUpButton(PropertyCondition popupButtonCondition)
        {
            try
            {
                AutomationElement PopUpAUElement = GetBrowserChildAUElement(popUpPropertyCondition, TreeScope.Children);
                if (PopUpAUElement != null)
                {
                    LogDebug("Finding browser popup button element...");
                    AutomationElement ButtonElement = PopUpAUElement.FindFirst(TreeScope.Descendants, popupButtonCondition);
                    LogDebug("Clicking OK Button...");
                    ClickOnButton(ButtonElement);
                    return true;
                }

                return false;
            }
            catch (Exception e)
            {
                LogError("Close Browser PopUp", e);
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