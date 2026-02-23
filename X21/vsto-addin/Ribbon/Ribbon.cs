using X21.Common.Data;
using X21.Common.Model;
using X21.Excel;
using X21.Utils;
using Microsoft.Office.Core;
using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using X21.Interfaces;
using X21.Logging;
using Office = Microsoft.Office.Core;
using System.Drawing;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace X21.Ribbon
{
    [ComVisible(true)]
    public class Ribbon : IRibbonExtensibility, IComponent
    {
        private IRibbonUI _ribbon;
        private ExcelSelection _excelSelection;

        public Container Container { get; }

        public void Init()
        {
            _excelSelection = Container.Resolve<ExcelSelection>();
            _excelSelection.SelectionChanged += OnSelectionChanged;
        }

        public void Exit()
        {
            _excelSelection.SelectionChanged -= OnSelectionChanged;
            _excelSelection = null;
        }

        //image mso: https://bert-toolkit.com/imagemso-list.html
        public Ribbon(Container container)
        {
            Container = container;
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("X21.Ribbon.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this._ribbon = ribbonUI;
        }

        private void OnSelectionChanged(object sender, Excel.Events.SelectionChangedEventArgs e)
        {
            _ribbon?.Invalidate();
        }

        public string OnGetLabel(IRibbonControl control)
        {
            return Execute.Call(
                () =>
                {
                    var cmd = Container.CommandById(control.Id);

                    var text = cmd?.Annotation.Title ?? string.Empty;

                    return RibbonXmlUtils.EscapeLabel(text);
                });
        }

        public string OnGetTabLabel(IRibbonControl control)
        {
#if DEBUG
            return "X21-debug";
#else
            return "X21";
#endif
        }

        public void OnAction(IRibbonControl control)
        {
            Execute.Call(
                () =>
                {
                    var cmd = Container.CommandById(control.Id);

                    if (cmd != null)
                    {
                        if (cmd.CanExecute(null))
                        {
                            cmd.Execute(null);
                        }
                        else
                        {
                            Logger.Info("WARNING: Command {0} can not be executed now!", control.Id);
                        }
                    }
                    else
                    {
                        Logger.Info("WARNING: Command {0} not found!", control.Id);
                    }
                });
        }

        public bool OnGetEnabled(IRibbonControl control)
        {
            return Execute.Call(
                () =>
                {
                    var cmd = Container.CommandById(control.Id);

                    if (cmd != null)
                    {
                        return cmd.CanExecute(null);
                    }

                    Logger.Info("WARNING: Command {0} not found!", control.Id);
                    return false;
                }
            );
        }

        public bool OnGetVisible(IRibbonControl control)
        {
            return Execute.Call(
                () =>
                {
                    // Only control visibility for CommandToggleTaskPaneHome
                    if (control.Id == "CommandToggleTaskPaneHome")
                    {
                        // Use RegistryHelper to get the visibility setting
                        // Default is true (visible) if not set
                        // GetBool handles both old string format and new DWORD format
                        bool isVisible = RegistryHelper.GetBool("ShowChatOnHomeRibbon", true);
                        Logger.Info($"CommandToggleTaskPaneHome visibility: {(isVisible ? "visible" : "hidden")}");
                        return isVisible;
                    }

                    // For all other controls, default to visible
                    return true;
                }
            );
        }

        public Bitmap OnGetImage(IRibbonControl control)
        {
            var assembly = Assembly.GetExecutingAssembly();

            // Get the appropriate icon name based on the control ID
            string iconName = control.Id switch
            {
                "CommandToggleTaskPane" => "x21_logo_80.png",
                "CommandToggleTaskPaneHome" => "x21_logo_80.png",
                _ => "x21_logo_80.png"
            };

            using (Stream stream = assembly.GetManifestResourceStream($"X21.Ribbon.Icons.{iconName}"))
            {
                if (stream == null)
                {
                    Logger.Info($"Failed to load {iconName} - stream is null");
                    return null;
                }
                return new Bitmap(stream);
            }
        }

        public string OnGetVersionLabel(IRibbonControl control)
        {
            return EnvironmentHelper.GetVersion();
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
