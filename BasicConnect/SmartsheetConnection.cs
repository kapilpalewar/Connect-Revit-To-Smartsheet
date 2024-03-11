using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.ServiceModel.Syndication;
using System.Windows.Forms;
using System.IO;

namespace BasicConnect
{
    public class SmartsheetConnection
    {
        public IList<(long folderId, string folderName)> GetFolderIdsAndNames()
        {
            string accessToken = "4XXHcMUhOBZ3RMDqZApNeacxoKp0n1Cg8m3IN";
            long workspaceId = 8232235724760964;

            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

            // Get the list of folders in the workspace
            PaginatedResult<Folder> foldersResult = smartsheet.WorkspaceResources.FolderResources.ListFolders(workspaceId, null);
             
            // Retrieve the folder IDs and names
            IList<(long folderId, string folderName)> folderIdsAndNames = new List<(long folderId, string folderName)>();
            foreach (Folder folder in foldersResult.Data)
            {
                folderIdsAndNames.Add(((long folderId, string folderName))(folder.Id, folder.Name));
            }
            return folderIdsAndNames;
        }
        public Sheet ConnectDB(long selectedSheetId)
        {
            string accessToken = "4XXHcMUhOBZ3RMDqZApNeacxoKp0n1Cg8m3IN";

            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

            Sheet sheet = smartsheet.SheetResources.GetSheet(selectedSheetId, null, null, null, null, null, null, null);

            if (sheet != null)
            {
                return sheet;
            }
            else
            {
                Console.WriteLine($"Sheet with ID '{selectedSheetId}' not found.");
                return null;
            }
        }
        public string RetrieveProjectNumber(Document document)
        {
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
        }

    }
}
