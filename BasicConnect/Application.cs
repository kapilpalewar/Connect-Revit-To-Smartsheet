using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using BasicConnect.Forms;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Media.Imaging;

namespace BasicConnect
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Application : IExternalApplication
    {
        public static DocumentSet Documents { get; internal set; }

        private PushButton button;
        private string thisAssemblyPath;
        private int mismatchCount = 0;

        public Result OnStartup(UIControlledApplication application)
        {
            // Register the event handler for document opening
            application.ControlledApplication.DocumentOpened += ControlledApplication_DocumentOpened;

            RibbonPanel panel = RibbonPanel(application);
            thisAssemblyPath = Assembly.GetExecutingAssembly().Location;

            button = panel.AddItem(new PushButtonData("BasicConnect", "BasicConnect", thisAssemblyPath, "BasicConnect.Command")) as PushButton;

            button.ToolTip = "Export data from Smartsheet to Revit";

            Uri uri = new Uri(Path.Combine(Path.GetDirectoryName(thisAssemblyPath), "Resources", "Base-4.ico"));
            BitmapImage bitmap = new BitmapImage(uri);
            button.LargeImage = bitmap;

            application.ApplicationClosing += a_ApplicationClosing;

            application.Idling += a_Idling;

            return Result.Succeeded;
        }

        public void ControlledApplication_DocumentOpened(object sender, DocumentOpenedEventArgs e)
        {
            Document doc = e.Document;
            ExecutePluginLogic(doc);

            // Update the button icon based on the mismatch count
            UpdateButtonIcon();
        }

        private void ExecutePluginLogic(Document doc)
        {
            // Add any additional plugin logic you require
            Mismatch mismatch = new Mismatch(doc);

            // Call the project number 
            mismatch.GetRevitProjectNumber();

            // Get all the sheets based on project number
            mismatchCount = mismatch.PopulateFolderIdAndMatchingData();

            // Display the mismatch count in a MessageBox
            //MessageBox.Show("Mismatch count: " + mismatchCount.ToString());
        }

        private void UpdateButtonIcon()
        {
            string iconName = "checked.ico";

            // Check the mismatch count and update the icon accordingly
            if (mismatchCount > 0)
            {
                // Update the icon to a red version
                //iconName = "Base-4_Red.ico";
                iconName = "circle.ico";
            }

            Uri uri = new Uri(Path.Combine(Path.GetDirectoryName(thisAssemblyPath), "Resources", iconName));
            BitmapImage bitmap = new BitmapImage(uri);
            button.LargeImage = bitmap;
        }

        private void a_Idling(object sender, IdlingEventArgs e)
        {
            //throw new NotImplementedException();
        }

        private void a_ApplicationClosing(object sender, ApplicationClosingEventArgs e)
        {
            //throw new NotImplementedException();
        }

        private RibbonPanel RibbonPanel(UIControlledApplication application)
        {
            string tab = "SmartSheet";
            RibbonPanel ribbonPanel = null;
            try
            {
                application.CreateRibbonTab(tab);
            }
            catch
            {
            }
            try
            {
                RibbonPanel panel = application.CreateRibbonPanel(tab, "test");
            }
            catch
            {
            }

            List<RibbonPanel> panels = application.GetRibbonPanels(tab);

            foreach (RibbonPanel p in panels)
            {
                if (p.Name == "test")
                {
                    ribbonPanel = p;
                }
            }
            return ribbonPanel;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}
