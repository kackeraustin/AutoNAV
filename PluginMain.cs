using System;
using Autodesk.Navisworks.Api.Plugins;

namespace AutoNAVSearchSets
{
    [PluginAttribute("AutoNAVSearchSets",
                     "ACLP_VDC",
                     ToolTip = "Automatic Search Set Creation for Navisworks",
                     DisplayName = "AutoNav SearchSets")]
    [AddInPluginAttribute(AddInLocation.AddIn)]
    public class PluginMain : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            try
            {
                MainWindow mainWindow = new MainWindow();
                mainWindow.ShowDialog();
                return 0;
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(
                    "Error: " + ex.Message,
                    "AutoNav SearchSets Error",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Error);
                return 1;
            }
        }
    }
}
