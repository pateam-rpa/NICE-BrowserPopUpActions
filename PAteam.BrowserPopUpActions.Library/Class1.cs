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
        private static readonly string PackageName = "Direct.BrowserDialogActions.Library";

        private static void LogDebug(string message)
        {
            if (_log.IsDebugEnabled)
            {
                _log.Debug(PackageName + " - " + message);
            }
        }

        private static void LogError(string message, Exception exception)
        {
            if (_log.IsErrorEnabled)
            {
                _log.Debug(PackageName + " - " + message, exception);
            }
        }

        [DirectDom("Get dialog existence")]
        [DirectDomMethod("Get dialog existence")]
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

        [DirectDom("Get dialog text")]
        [DirectDomMethod("Get dialog text")]
        [MethodDescription("Returns browser Dialog text on active tab if found")]
        public static string GetDialogText()
        {
            try
            {
                AutomationElement DialogAUElement = GetDialogElement();
                if (DialogAUElement != null)
                {
                    Condition LabelPropertyCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text);
                    LogDebug("Finding browser dialog text elements...");
                    AutomationElementCollection TextElements = DialogAUElement.FindAll(TreeScope.Descendants, LabelPropertyCondition);
                    StringBuilder finalText = new StringBuilder();
                    LogDebug("Found " + TextElements.Count.ToString() + " text elements");
                    foreach (AutomationElement TextElement in TextElements)
                    {
                        string line = GetAUElementText(TextElement);
                        LogDebug("Retrieving text property from found text element");
                        if (!string.IsNullOrEmpty(line))
                        {
                            finalText.AppendLine(GetAUElementText(TextElement));
                        }
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

        [DirectDom("Open file dialog")]
        [DirectDomMethod("Open file dialog")]
        [MethodDescription("Opens File Dialog to select files to upload")]
        public static bool OpenFileDialog()
        {
            try
            {
                return ClickDialogButtonByCondition(new PropertyCondition(AutomationElement.AutomationIdProperty, "file-upload-button"));
            }
            catch (Exception e)
            {
                LogError("Failed to open file dialog", e);
                return false;
            }
        }

        [DirectDom("Click dialog button")]
        [DirectDomMethod("Click dialog button {button name}")]
        [MethodDescription("Clicks on dialog button with specified name")]
        public static bool ClickDialogButtonByName(string buttonName)
        {
            if (string.IsNullOrEmpty(buttonName))
            {
                LogDebug("Button name cannot be empty");
                return false;
            }

            try
            {
                Condition condition = new AndCondition(
                    new PropertyCondition(AutomationElement.NameProperty, buttonName),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                );
                return ClickDialogButton(condition);
            }
            catch (Exception e)
            {
                LogError("Failed to click on a button: " + buttonName, e);
                return false;
            }

        }

        [DirectDom("Set dialog edit value")]
        [DirectDomMethod("Set dialog edit {edit name} value {value}")]
        [MethodDescription("Sets dialog edit value")]
        public static bool SetDialogEditValue(string editName, string value)
        {
            try
            {
                if (string.IsNullOrEmpty(editName))
                {
                    LogDebug("Edit name cannot be empty");
                    return false;
                }

                AutomationElement dialogElement = GetDialogElement();
                if (dialogElement == null)
                {
                    return false;
                }

                Condition editElementCondition = new AndCondition(
                    new PropertyCondition(AutomationElement.NameProperty, editName),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit)
                );

                AutomationElement editElement = dialogElement.FindFirst(TreeScope.Descendants, editElementCondition);

                if (editElement == null)
                {
                    LogDebug("Failed to find edit element " + editName);
                    return false;
                }

                if (!editElement.Current.IsEnabled)
                {
                    LogDebug("Edit element: " + editName + " is not enabled");
                    return false;
                }

                object valuePattern = null;

                if (editElement.TryGetCurrentPattern(ValuePattern.Pattern, out valuePattern))
                {
                    ((ValuePattern)valuePattern).SetValue(value);
                    return true;
                }
                else
                {
                    LogDebug("Edit element: " + editName + " does not support value pattern");
                    return false;
                }
            }
            catch (Exception e)
            {
                LogError("Failed to set edit: " + editName + " value", e);
                return false;
            }

        }

        private static AutomationElement GetDialogElement()
        {
            LogDebug("Creating localized conditions, only following langugaes are supported: PL, EN, DE, IT, SP");
            Condition[] dialogLocalizedConditions =
            {
                new PropertyCondition(AutomationElement.LocalizedControlTypeProperty, "Dialogfeld"),
                new PropertyCondition(AutomationElement.LocalizedControlTypeProperty, "dialog"),
                new PropertyCondition(AutomationElement.LocalizedControlTypeProperty, "okno dialogowe"),
                new PropertyCondition(AutomationElement.LocalizedControlTypeProperty, "diálogo"),
                new PropertyCondition(AutomationElement.LocalizedControlTypeProperty, "dialogo"),
            };
            Condition DialogPropertyCondition = new OrCondition(dialogLocalizedConditions);

            return GetBrowserChildAUElement(DialogPropertyCondition, TreeScope.Children);
        }

        private static bool ClickDialogButtonByCondition(Condition providedCondition)
        {
            Condition[] conditions = { new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button), providedCondition };

            Condition condition = new AndCondition(conditions);

            return ClickDialogButton(condition);
        }

        private static bool GetDialogExistence()
        {
            LogDebug("Finding browser Dialog element...");
            AutomationElement DialogAUElement = GetDialogElement();

            if (DialogAUElement != null)
            {
                LogDebug("Found dialog with localized control type: " + DialogAUElement.Current.LocalizedControlType + " and name: " + DialogAUElement.Current.Name);
                return true;
            }

            return false;
        }

        private static AutomationElement GetBrowserChildAUElement(Condition elementCondition, TreeScope scope)
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

        private static bool ClickDialogButton(Condition DialogButtonCondition)
        {
            try
            {
                AutomationElement DialogAUElement = GetDialogElement();
                if (DialogAUElement != null)
                {
                    LogDebug("Finding browser Dialog button element...");
                    AutomationElement ButtonElement = DialogAUElement.FindFirst(TreeScope.Descendants, DialogButtonCondition);
                    LogDebug("Clicking button...");
                    ClickOnButton(ButtonElement);
                    return true;
                }

                return false;
            }
            catch (Exception e)
            {
                LogError("Failed to click on a dialog button. Error: ", e);
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