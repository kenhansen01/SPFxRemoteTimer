using OfficeDevPnP.Core.Utilities;
using System.Net;

namespace EmployeesRemoteTimer
{
    class Program
    {
        static void Main(string[] args)
        {

            EmployeeJob simpleJob = new EmployeeJob();

            simpleJob.UseNetworkCredentialsAuthentication("sp_farm", "password", "CORP");

            simpleJob.AddSite("http://dev-sharepoint/sites/usermanagement");
            simpleJob.AddSite("http://sharepoint/sites/usermanagement");

            simpleJob.DataHubCred = new NetworkCredential("DataAPIUser", "password", "corp"); //CredentialManager.GetCredential("datahub");
            simpleJob.DataHubBaseUrl = "https://someplace.corp.com:8443/";
            simpleJob.DataHubRequestUri = "hrdatahub/v1/employee";
            simpleJob.ListName = "Employees";
            
            foreach(string arg in args)
            {
                switch (arg.ToUpper())
                {
                    case "/ALL":
                        simpleJob.RunFullAudit = true;
                        break;
                    default:
                        simpleJob.RunFullAudit = false;
                        break;
                }
            }

            simpleJob.Run();

        }
    }
}
