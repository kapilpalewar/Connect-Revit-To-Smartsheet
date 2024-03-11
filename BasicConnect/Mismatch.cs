using Autodesk.Revit.DB;
using Smartsheet.Api.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.UI;
using Newtonsoft.Json;
using Smartsheet.Api;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Windows.Controls;

namespace BasicConnect
{
    public class Mismatch
    {
        private Document document;

        public Mismatch(Document doc)
        {
            document = doc;
        }

        public string GetRevitProjectNumber()
        {
            // Retrieve the Project Number from the Revit model
            // Get the project information
            ProjectInfo projectInfo = document.ProjectInformation;

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




        // Create a dictionary to store the folder name and IDs
        // Create a dictionary to store the folder name and IDs
        Dictionary<string, long> folderIdNameMap = new Dictionary<string, long>();
        Dictionary<string, Dictionary<string, string>> matchingData = new Dictionary<string, Dictionary<string, string>>();

        public int PopulateFolderIdAndMatchingData()
        {
            // Get Smartsheet details
            string accessToken = "4XXHcMUhOBZ3RMDqZApNeacxoKp0n1Cg8m3IN";
            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

            // Get the Project Number from the Revit model
            string projectNumber = GetRevitProjectNumber();

            if (string.IsNullOrEmpty(projectNumber))
            {
                MessageBox.Show("Project Number not found in the Revit model.", "Error");
                return 0;
            }

            // Clear existing items in folderIdNameMap and matchingData
            folderIdNameMap.Clear();
            matchingData.Clear();
            int mismatchCount = 0;

            // Retrieve folders and sheets
            SmartsheetConnection smartsheetConnection = new SmartsheetConnection();
            IList<(long folderId, string folderName)> folderIdsAndNames = smartsheetConnection.GetFolderIdsAndNames();

            try
            {
                // Populate the folderIdNameMap and matchingData
                foreach (var (folderId, folderName) in folderIdsAndNames)
                {
                    // Assuming the folder ID needs to be converted to a long value
                    long folderIdLong = Convert.ToInt64(folderId);

                    // Add the folder ID and corresponding name to the folderIdNameMap
                    folderIdNameMap[folderName] = folderIdLong;

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
                            // Retrieve the sheet object
                            Sheet sheet1 = smartsheet.SheetResources.GetSheet(sheetId, null, null, null, null, null, null, null);
                            if (sheet1 == null)
                            {
                                MessageBox.Show("Sheet is null.", "Error");
                                return 0;
                            }

                            // Retrieve the columns
                            List<Column> columns = sheet1.Columns
                                .Where(c => c.Title == "Testing" || c.Title == "CONTENTS")
                                .ToList();

                            // Retrieve the rows for the selected columns
                            Dictionary<string, string> matchingRowData = new Dictionary<string, string>();

                            foreach (Row row in sheet1.Rows)
                            {
                                string parameterName = string.Empty;
                                string parameterValue = string.Empty;

                                // Find the cell in the "Testing" column for Revit parameter check
                                Cell testingCell = row.Cells.FirstOrDefault(c => c.ColumnId == columns.FirstOrDefault(col => col.Title == "Testing")?.Id);
                                if (testingCell != null)
                                {
                                    parameterName = testingCell.DisplayValue;
                                }

                                // Find the cell in the "CONTENTS" column for Revit parameter value check
                                Cell contentsCell = row.Cells.FirstOrDefault(c => c.ColumnId == columns.FirstOrDefault(col => col.Title == "CONTENTS")?.Id);
                                if (contentsCell != null)
                                {
                                    parameterValue = contentsCell.DisplayValue;
                                }

                                // Check if Revit parameter name is present in the "Testing" column
                                if (!string.IsNullOrEmpty(parameterName) && !string.IsNullOrEmpty(parameterValue))
                                {
                                    // Compare Revit parameter value with cell value in the "CONTENTS" column
                                    // Adjust the comparison logic as per your requirement
                                    if (parameterValue != GetRevitParameterValue(parameterName))
                                    {
                                        matchingRowData[parameterName] = parameterValue;
                                        mismatchCount++;
                                    }
                                }
                            }

                            matchingData[sheetName] = matchingRowData;
                        }
                    }
                }

                // Return the mismatch count
                return mismatchCount;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error occurred while retrieving and populating folder IDs and matching data: " + ex.Message, "Error");
                return 0;
            }
        }


        private string GetRevitParameterValue(string parameterName)
        {
            // Example implementation:
            // Assuming you have access to the Revit Document object, you can use it to retrieve the parameter value
            // Replace "document" with your actual Document object

            // Get the Revit project parameters
            ProjectInfo projectInfo = document.ProjectInformation;
            ParameterSet parameters = projectInfo.Parameters;

            // Find the parameter by name
            foreach (Parameter parameter in parameters)
            {
                if (parameter.Definition.Name == parameterName)
                {
                    // Check the parameter type and retrieve the value accordingly
                    switch (parameter.StorageType)
                    {
                        case StorageType.String:
                            return parameter.AsString();

                        case StorageType.Integer:
                            return parameter.AsInteger().ToString();

                        case StorageType.Double:
                            return parameter.AsDouble().ToString();

                        // Handle other parameter types as needed

                        default:
                            return string.Empty;
                    }
                }
            }

            return string.Empty;
        }
    }
}
