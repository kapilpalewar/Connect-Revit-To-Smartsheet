using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Newtonsoft.Json;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;


namespace BasicConnect.Forms
{
    public partial class XYZForm : System.Windows.Forms.Form
    {
        private UIApplication uiapp;
        private UIDocument uidoc;
        private Autodesk.Revit.ApplicationServices.Application app;
        private Document doc;
        private object commandData;

       
       

        public XYZForm(ExternalCommandData commandData)
        {
            InitializeComponent();
            uiapp = commandData.Application;
            uidoc = uiapp.ActiveUIDocument;
            app = uiapp.Application;
            doc = uidoc.Document;


        }

        public XYZForm(Document doc)
        {
            this.doc = doc;
            
        }

        private void UpdateTextBox(IList<Column> columns, IList<Row> rows)
        {

            dataGridView1.Columns.Clear();

            // Append column names to the TextBox
            foreach (Column column in columns)
            {
                dataGridView1.Columns.Add(column.Title, column.Title);
            }
           
            // Loop through the rows and append the data to the TextBox
            foreach (Row row in rows)
            {
                object[] rowValues = new object[columns.Count];
                int i = 0;
                foreach (Cell cell in row.Cells)
                {
                    rowValues[i] = cell.Value;
                    i++;
                }
                dataGridView1.Rows.Add(rowValues);


            }
        }            



        //Create Dic to store the folder name and ids
        private Dictionary<string, long> folderIdNameMap = new Dictionary<string, long>();

        // Get the sheet by project number
        private Dictionary<string, long> sheetIdNameMap1 = new Dictionary<string, long>();

        private void XYZForm_Load(object sender, EventArgs e)
        {
            //Get all folders from the sheets and send to CheckListBox
            SmartsheetConnection smartsheetConnection = new SmartsheetConnection();
            IList<(long folderId, string folderName)> folderIdsAndNames = smartsheetConnection.GetFolderIdsAndNames();
            try
            {            

                // Clear the folder ID and name map
                folderIdNameMap.Clear();

                // Populate the folder IDs and names map
                foreach (var (folderId, folderName) in folderIdsAndNames)
                {
                    // Assuming the folder ID needs to be converted to a long value
                    long folderIdLong = Convert.ToInt64(folderId);

                    // Add the folder ID and corresponding name to the dictionary
                    folderIdNameMap[folderName] = folderIdLong;

                   
                }

                // MessageBox.Show("Folder IDs and names successfully retrieved and populated in the CheckedListBox.", "Success");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error occurred while retrieving and populating folder IDs and names: " + ex.Message, "Error");
            }

            // Get the sheet by project number             
            string accessToken = "4XXHcMUhOBZ3RMDqZApNeacxoKp0n1Cg8m3IN";
            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

            // Get the Project Number from the Revit model
            string projectNumber = GetRevitProjectNumber();

            if (string.IsNullOrEmpty(projectNumber))
            {
                MessageBox.Show("Project Number not found in the Revit model.", "Error");
                return;
            }

            // Clear existing items in listBox1 and sheetIdNameMap           
            sheetIdNameMap1.Clear();
            // Flag to track if any sheet matches the project number
            bool projectNumberMatch = false;

            // Retrieve sheet names and IDs from all folders in folderIdNameMap
            foreach (var folderEntry in folderIdNameMap)
            {
                string folderName = folderEntry.Key;
                long folderId = folderEntry.Value;

                // Get the folder object
                Folder fullFolder = smartsheet.FolderResources.GetFolder(folderId, null);

                // Retrieve the sheet names and IDs
                foreach (Sheet sheet in fullFolder.Sheets)
                {
                    string sheetName = sheet.Name;
                    long sheetId = (long)sheet.Id;

                    // Check if the sheet name contains the Project Number
                    if (sheetName.Contains(projectNumber))
                    {                      
                        label2.Text = sheetName;
                        sheetIdNameMap1[sheetName] = sheetId;
                        projectNumberMatch = true;
                        Sheet sheet1 = smartsheet.SheetResources.GetSheet(sheetId, null, null, null, null, null, null, null);
                        if (sheet1 == null)
                        {
                            MessageBox.Show("Sheet is null.", "Error");
                            return;
                        }

                        IList<Column> columns = sheet1.Columns;
                        IList<Row> rows = sheet1.Rows;

                        UpdateTextBox(columns, rows);
                    }
                }
            }
            // Display the appropriate MessageBox based on the projectNumberMatch flag
            if (projectNumberMatch)
            {
                //MessageBox.Show("Sheets filtered and retrieved successfully.", "Success");
            }
            else
            {
                MessageBox.Show("Project Number does not match with any sheets.", "Warning");
            }

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            //{
            //    // Get the selected row
            //    DataGridViewRow selectedRow = dataGridView1.Rows[e.RowIndex];

            //    // Get the value of a specific cell in the selected row
            //    string cellValue = selectedRow.Cells["Column2"].Value.ToString();

            //    using (Transaction transaction = new Transaction(doc, "Update Project Name"))
            //    {
            //        transaction.Start();

            //        // Set the project name to the value of the selected cell in the dataGridView1 control
            //        ProjectInfo projectInfo = doc.ProjectInformation;
            //        //DataGridViewRow selectedRow = dataGridView1.CurrentRow;
            //        if (selectedRow != null)
            //        {
            //            DataGridViewCell cell = selectedRow.Cells["Column2"];
            //            if (cell != null && cell.Value != null)
            //            {
            //                projectInfo.Name = cell.Value.ToString();
            //            }
            //        }
            //        TaskDialog.Show("Export Data", "Values added Sucessfully !!!");
            //        transaction.Commit();
            //    }
            //}
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }      

        //Get Sheet button  
        private void button2_Click(object sender, EventArgs e)
        {
           


        }
           

        private string GetRevitProjectNumber()
        {
            // Retrieve the Project Number from the Revit model
            // Get the project information
            ProjectInfo projectInfo = doc.ProjectInformation;

            if (projectInfo != null)
            {
                // Retrieve the project number
                Parameter projectNumberParameter = projectInfo.get_Parameter(BuiltInParameter.PROJECT_NUMBER);

                if (projectNumberParameter != null && projectNumberParameter.HasValue)
                {
                    // Return the project number value
                    return projectNumberParameter.AsString();
                }
            }

            // Project number not found or empty
            return string.Empty;
            //// Implement the necessary logic to extract the Project Number from the Revit model
            //// Replace this placeholder code with your actual implementation
            //string projectNumber = "B4-260-2201"; // Placeholder value, replace with actual logic
            //return projectNumber;
        }
        

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //try
            //{
            //    string selectedSheetName = listBox1.SelectedItem?.ToString();

            //    if (string.IsNullOrEmpty(selectedSheetName))
            //    {
            //        MessageBox.Show("No sheet selected.", "Error");
            //        return;
            //    }

            //    SmartsheetConnection abc = new SmartsheetConnection();

            //    // Retrieve the sheet ID from the dictionary
            //    if (sheetIdNameMap1.TryGetValue(selectedSheetName, out long sheetId))
            //    {
            //        Sheet sheet = abc.ConnectDB(sheetId);

            //        if (sheet == null)
            //        {
            //            MessageBox.Show("Sheet is null.", "Error");
            //            return;
            //        }

            //        IList<Column> columns = sheet.Columns;
            //        IList<Row> rows = sheet.Rows;

            //        UpdateTextBox(columns, rows);
            //    }
            //    else
            //    {
            //        MessageBox.Show($"Sheet ID not found for the selected sheet name: {selectedSheetName}", "Error");
            //    }
            //}
            //catch (NullReferenceException ex)
            //{
            //    MessageBox.Show("NullReferenceException occurred:\n" + ex.Message, "Error");
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("An error occurred:\n" + ex.Message, "Error");
            //}
        }

        // Show Mismatch parameters values
        // Show Mismatch parameters values
        private void button3_Click(object sender, EventArgs e)
        {
            // Get the project information element
            ElementId projectId = doc.ProjectInformation.Id;
            Element projectInfoElement = doc.GetElement(projectId);

            // Retrieve the parameters of the project information element
            ParameterSet parameterSet = projectInfoElement.Parameters;

            // Create a dictionary to store parameter names and corresponding values from DataGridView
            Dictionary<string, string> parameterValues = new Dictionary<string, string>();

            // Loop through each parameter and retrieve its name
            foreach (Parameter param in parameterSet)
            {
                string paramName = param.Definition.Name;
                string cellValue = GetCellValueFromColumn2(paramName);
                if (!string.IsNullOrEmpty(cellValue))
                {
                    parameterValues[paramName] = cellValue;
                }
            }

            // Create a list to store parameters with mismatching values
            List<string> mismatchedParameters = new List<string>();

            // Check for mismatching values
            foreach (KeyValuePair<string, string> entry in parameterValues)
            {
                Parameter parameter = projectInfoElement.LookupParameter(entry.Key);
                if (parameter != null)
                {
                    string parameterValue = parameter.AsString();
                    if (!entry.Value.Equals(parameterValue, StringComparison.OrdinalIgnoreCase))
                    {
                        mismatchedParameters.Add(entry.Key);
                    }
                }
            }

            if (mismatchedParameters.Count > 0)
            {
                //string mismatchedParamsString = string.Join(", ", mismatchedParameters);
                //MessageBox.Show("Mismatching parameters: " + mismatchedParamsString);
            }
            else
            {
                MessageBox.Show("All parameters have matching values.");
            }

            // Create a DataTable to store the data for DataGridView
            DataTable dt = new DataTable();
            dt.Columns.Add("Parameter");
            dt.Columns.Add("Revit Value");
            dt.Columns.Add("Sheet Value");
            dt.Columns.Add("Modified By");

            // Loop through each mismatched parameter and retrieve its Revit, sheet, and modified by values
            foreach (string mismatchedParameter in mismatchedParameters)
            {
                Parameter parameter = projectInfoElement.LookupParameter(mismatchedParameter);
                string revitValue = parameter != null ? parameter.AsString() : string.Empty;
                string sheetValue = parameterValues[mismatchedParameter];
                string modifiedBy = GetModifiedByValue(mismatchedParameter);

                DataRow row = dt.NewRow();
                row["Parameter"] = mismatchedParameter;
                row["Revit Value"] = revitValue;
                row["Sheet Value"] = sheetValue;
                row["Modified By"] = modifiedBy;
                dt.Rows.Add(row);
            }
            // Bind the DataTable to the DataGridView
            dataGridView2.DataSource = dt;
        }
        
        // Helper method to get cell value from DataGridView based on matching parameter name
        private string GetCellValueFromColumn2(string paramName)
        {
            string cellValue = string.Empty;

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells["Testing"].Value != null)
                {
                    string parameterName = row.Cells["Testing"].Value.ToString();
                    if (parameterName.Equals(paramName, StringComparison.OrdinalIgnoreCase))
                    {
                        if (row.Cells["CONTENTS"].Value != null)
                        {
                            cellValue = row.Cells["CONTENTS"].Value.ToString();
                        }
                        break;
                    }
                }
            }

            return cellValue;
        }

        // Method to retrieve the "Modified By" value from the sheets based on the parameter name
        private string GetModifiedByValue(string paramName)
        {
            string modifiedBy = string.Empty;

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells["Testing"].Value != null)
                {
                    string parameterName = row.Cells["Testing"].Value.ToString();
                    if (parameterName.Equals(paramName, StringComparison.OrdinalIgnoreCase))
                    {
                        if (row.Cells["Modified By"].Value != null)
                        {
                            modifiedBy = row.Cells["Modified By"].Value.ToString();
                        }
                        break;
                    }
                }
            }
            return modifiedBy;
        }


        private void button4_Click(object sender, EventArgs e)
        {

            // Get the project information element
            ElementId projectId = doc.ProjectInformation.Id;
            Element projectInfoElement = doc.GetElement(projectId);

            // Update the project parameters
            using (Transaction transaction = new Transaction(doc, "Update Project Parameters"))
            {
                transaction.Start();

                // Loop through the selected rows in the DataGridView
                foreach (DataGridViewRow row in dataGridView2.SelectedRows)
                {
                    string parameterName = row.Cells["Parameter"].Value.ToString();
                    string parameterValue = row.Cells["Sheet Value"].Value.ToString();

                    // Update the project parameter with the new value
                    Parameter parameter = projectInfoElement.LookupParameter(parameterName);
                    if (parameter != null)
                    {
                        parameter.Set(parameterValue);
                    }
                }

                transaction.Commit();
            }

            MessageBox.Show("Project parameters updated successfully!");
        }
    }
}
