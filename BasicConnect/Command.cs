using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using BasicConnect.Forms;
using Smartsheet.Api.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BasicConnect
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            XYZForm form1 = new XYZForm(commandData);
            form1.ShowDialog();

            using (Transaction tx = new Transaction(doc))
            {
                try
                {
                    tx.Start("Button1_Click");

                    TaskDialog td = new TaskDialog("This is first connection");

                    // Add your code logic here...

                    tx.Commit();
                }
                catch (Exception e)
                {
                    TaskDialog td = new TaskDialog("Error");

                    td.Title = "Error";
                    td.AllowCancellation = true;
                    td.MainInstruction = "Error Found";
                    td.MainContent = "Something went wrong: " + e.Message;
                    td.CommonButtons = TaskDialogCommonButtons.Close;

                    td.Show();

                    Debug.Print(e.Message);
                    tx.RollBack();
                }
            }

            return Result.Succeeded;
        }
    }
}
