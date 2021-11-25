using Direct.Shared;
using System;
using log4net;
using System.Windows.Automation;

namespace Direct.BrowserPopUpActions.Library
{
    [DirectSealed]
    [DirectDom("Browser PopUp Functions")]
    [ParameterType(false)]
    public static class BrowserPopUpFunctions
    {
        private static readonly ILog _log = LogManager.GetLogger("LibraryObjects");
        private static string packageName = "Direct.BrowserPopUpActions.Library";

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

        [DirectDom("Get PopUp Existence")]
        [DirectDomMethod("Get Browser PopUp Existence")]
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

        [DirectDom("Get PopUp Text")]
        [DirectDomMethod("Get Browser PopUp Text")]
        [MethodDescription("Iterates over all browser windows and return text from popup window if found")]
        public static string GetPopUpText()
        {
            string popUpText = "";
            try
            {
                AutomationElement PopUpAUElement = GetPopUpAUElement();
                if (PopUpAUElement != null)
                {
                    Condition LabelPropertyCondition = new PropertyCondition(AutomationElement.LocalizedControlTypeProperty, "text");
                    LogDebug("Finding browser popup text elements...");
                    AutomationElementCollection TextElements = PopUpAUElement.FindAll(TreeScope.Descendants, LabelPropertyCondition);
                    LogDebug("Found " + TextElements.Count.ToString() + " text elements");
                    foreach (AutomationElement TextElement in TextElements)
                    {
                        LogDebug("Retrieving text property from found text element");
                        popUpText = popUpText + GetAUElementText(TextElement) + "$$$";

                    }

                    return popUpText;
                }

                return string.Empty;
            }
            catch (Exception e)
            {
                LogError("Get Browser PopUp Text Exception", e);
                return string.Empty;
            }
        }

        [DirectDom("Close Browser PopUp")]
        [DirectDomMethod("Close Browser PopUp")]
        [MethodDescription("Clicks OK Button on found browser pop up")]
        public static bool ClosePopUp()
        {
            try
            {
                AutomationElement PopUpAUElement = GetPopUpAUElement();
                if (PopUpAUElement != null)
                {
                    Condition ButtonPropertyCondition = new PropertyCondition(AutomationElement.LocalizedControlTypeProperty, "button");
                    LogDebug("Finding browser popup OK button element...");
                    AutomationElement ButtonElement = PopUpAUElement.FindFirst(TreeScope.Descendants, ButtonPropertyCondition);
                    LogDebug("Clicking OK Button...");
                    object invokePattern = null;
                    ButtonElement.TryGetCurrentPattern(InvokePattern.Pattern, out invokePattern);
                    ((InvokePattern)invokePattern).Invoke();
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

        private static bool GetPopUpExistence()
        {
            LogDebug("Finding browser popup element...");
            AutomationElement PopUpAUElement = GetPopUpAUElement();

            if (PopUpAUElement != null)
            {
                LogDebug("Found dialog with localized control type: " + PopUpAUElement.Current.LocalizedControlType + " and name: " + PopUpAUElement.Current.Name);
                return true;
            }

            return false;
        }

        private static AutomationElement GetPopUpAUElement()
        {
            AutomationElementCollection BrowserAUElements = GetBrowserAUElements();
            if (BrowserAUElements.Count == 0)
            {
                LogDebug("Browser elements not found! Please check if browser is open and active");
                return null;
            }

            LogDebug("BrowserAUElements Count: " + BrowserAUElements.Count.ToString());
            Condition PopUpPropertyCondition = new PropertyCondition(AutomationElement.LocalizedControlTypeProperty, "dialog");
            Condition BrowserRootPropertyCondition = new PropertyCondition(AutomationElement.ClassNameProperty, "BrowserRootView");
            foreach (AutomationElement BrowserAUElement in BrowserAUElements)
            {
                AutomationElement BrowserRootAUElement = BrowserAUElement.FindFirst(TreeScope.Children, BrowserRootPropertyCondition);
                if (BrowserRootAUElement != null)
                {
                    AutomationElement PopUpAUElement = BrowserRootAUElement.FindFirst(TreeScope.Children, PopUpPropertyCondition);
                    if (PopUpAUElement != null)
                    {
                        return PopUpAUElement;
                    }
                }
            }

            LogDebug("Browser popup element not found. If this is incorrect, check if page with popup is acitve or try reloading and activate target page.");
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