using DataHubAPIClient.DataHubAPIRequests;
using DataHubAPIClient.DataHubAPIResponse;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.TimerJobs;
using System;
using System.Collections.Specialized;
using System.Linq;
using System.Net;
using System.Reflection;

namespace EmployeesRemoteTimer
{
    public class EmployeeJob : TimerJob
    {

        #region Internal Variables
        protected internal PropertyInfo[] employeeProps = typeof(Employee).GetProperties().Where(p => p.Name != "Department").ToArray();
        protected internal PropertyInfo[] departmentProps = typeof(Department).GetProperties();
        #endregion

        #region Private Variables

        // DataHub Information
        private NetworkCredential dataHubCred;
        private string dataHubBaseUrl;
        private string dataHubRequestUri;
        private StringCollection rqParams;
        private DataHubRequests hubRequests;

        // TimerState Information
        private bool runFullAudit = false;
        private string previousRun;

        // Scope variables
        private string listName;
        #endregion

        #region Constructor
        /// <summary>
        /// EmployeeJob Constructor. Sets ManageState to true and adds EmployeeJob_TimerJobRun to be run.
        /// </summary>
        public EmployeeJob() : base("EmployeeJob")
        {
            TimerJobRun += EmployeeJob_TimerJobRun;
            ManageState = true;
        }

        #endregion

        #region Job information and state management
        /// <summary>
        /// Sets the <see cref="dataHubCred"/>
        /// </summary>
        public NetworkCredential DataHubCred { set => dataHubCred = value; }
        /// <summary>
        /// Sets the <see cref="dataHubBaseUrl"/>
        /// </summary>
        public string DataHubBaseUrl { set => dataHubBaseUrl = value; }
        /// <summary>
        /// Sets the <see cref="dataHubRequestUri"/>
        /// </summary>
        public string DataHubRequestUri { set => dataHubRequestUri = value; }
        /// <summary>
        /// Sets the <see cref="rqParams"/>
        /// </summary>
        public StringCollection RqParams { set => rqParams = value; }
        /// <summary>
        /// Sets <see cref="hubRequests"/>
        /// </summary>
        public DataHubRequests HubRequests { set => hubRequests = value; }        
        /// <summary>
        /// Sets <see cref="runFullAudit"/>
        /// </summary>
        public bool RunFullAudit { set => runFullAudit = value; }
        /// <summary>
        /// Sets <see cref="previousRun"/>
        /// </summary>
        public string PreviousRun { set => previousRun = value; }
        /// <summary>
        /// Sets <see cref="listName"/>
        /// </summary>
        public string ListName { get => listName; set => listName = value; }
        #endregion

        #region Run Job
        void EmployeeJob_TimerJobRun(object sender, TimerJobRunEventArgs e)
        {
            try
            {                
                HubRequests = new DataHubRequests(dataHubBaseUrl, dataHubRequestUri, dataHubCred.UserName, dataHubCred.Password);
                PreviousRun = e.PreviousRun.GetValueOrDefault().ToLocalTime().ToString("o");
                Log.Info("EmployeeJob :: Properties", "Check for custom properties on state object. Create if needed.");
                if (e.GetProperty("WeeklyAuditRun") == "")
                {
                    Log.Info("EmployeeJob :: Properties", "Create WeeklyAuditRun Property.");
                    e.SetProperty("WeeklyAuditRun", (new DateTime(2018, 2, 4, 0, 0, 0, DateTimeKind.Local)).ToString());
                }


                if (runFullAudit || DateTime.Now - DateTime.Parse(e.GetProperty("WeeklyAuditRun")) >= new TimeSpan(1, 0, 0, 0))
                {
                    // Run full audit to get all users from datahub and match to list
                    RqParams = new StringCollection() { "department.functionCode=060" };
                    FullAudit(e.WebClientContext);
                    if (DateTime.Now - DateTime.Parse(e.GetProperty("WeeklyAuditRun")) > new TimeSpan(1, 0, 0, 0))
                    {
                        e.SetProperty("WeeklyAuditRun", DateTime.Today.ToString());
                    }
                }
                else
                {
                    // Check for updates since last run.
                    RqParams = new StringCollection() { "department.functionCode=060" };
                    ChangeSinceLastRun(e.WebClientContext);
                }

                
                e.CurrentRunSuccessful = true;
                e.DeleteProperty("LastError");
            }
            catch(Exception ex)
            {
                Log.Error("EmployeeJob", "Error while processing web {0}. Error = {1}. Stack = {2}", e.Url, ex.Message, ex.StackTrace);
                e.CurrentRunSuccessful = false;
                e.SetProperty("LastError", ex.Message);
            }
        }

        private void JoinedOrLeft(DataHubResponse idsOnlyResponse, ListItem[] employeesInList, ClientContext ctx, List employeeList)
        {
            ListItem[] notIn = employeesInList
                .Where(eItem => idsOnlyResponse.Employees
                  .Select(eIds => eIds.EmployeeId)
                  .Contains(eItem["Title"].ToString()) == false)
                .ToArray();

            foreach (ListItem listItemEmployee in notIn)
            {
                DataHubResponse employeeResponse = hubRequests.RunQueryAsync((new StringCollection() { $"employeeId={listItemEmployee["Title"].ToString()}" })).Result;
                if (employeeResponse.RecordCount == 0)
                {
                    Log.Info("EmployeeJob :: Left", "Change status for {0} to \"Terminated\". No longer in DataHub.", listItemEmployee["Name"]);
                    listItemEmployee["Status"] = "Terminated";
                    listItemEmployee.Update();
                }
                else
                {
                    Log.Info("EmployeeJob :: Left", "Change status for {0} to \"Left \". No longer in .", listItemEmployee["Name"]);
                    Employee dataHubEmployee = employeeResponse.Employees[0];
                    dataHubEmployee.Status = "Left ";
                    UpdateEmployeeItemValues(listItemEmployee, dataHubEmployee);
                }
                ctx.ExecuteQueryRetry();
            }

            string[] newEmployeeIdsToAdd = idsOnlyResponse.Employees.Where(dEmp => employeesInList.Select(lEmp => lEmp["Title"].ToString()).Contains(dEmp.EmployeeId) == false).ToArray().Select(e => e.EmployeeId).ToArray();
            if(newEmployeeIdsToAdd.Length > 0)
            {
                DataHubResponse newEmployeesResponse = hubRequests.RunQueryAsync((new StringCollection() { $"employeeId=inList::{String.Join(",", newEmployeeIdsToAdd)}" })).Result;

                foreach (Employee dataHubEmployee in newEmployeesResponse.Employees)
                {
                    Log.Info("EmployeeJob :: Joined", "Add {0} to EmployeeList.", dataHubEmployee.Name);
                    AddEmployeeToList(employeeList, dataHubEmployee);
                    ctx.ExecuteQueryRetry();
                }
            }           

        }

        private void FullAudit(ClientContext ctx)
        {
            Log.Info("EmployeeJob :: FullAudit", "Compare SP List: {0}, to DataHub responses.", listName);

            List employeeList = ctx.Web.GetListByTitle(listName);

            ListItem[] employeeItems = GetEmployeesFromList(ctx, employeeList)
                .Where(eItem => eItem["Status"] == null || (eItem["Status"].ToString() != "Terminated" && eItem["Status"].ToString() != "Left ")).ToArray();
            DataHubResponse employeeIdsResponse = hubRequests.GetEmployeesIdsOnlyAsync(rqParams).Result;

            JoinedOrLeft(employeeIdsResponse, employeeItems, ctx, employeeList);

            RqParams = new StringCollection() { "department.functionCode=060" };

            DataHubResponse datahubEmployeesResponse = hubRequests.RunQueryAsync(rqParams).Result;

            foreach (Employee dataHubEmployee in datahubEmployeesResponse.Employees)
            {
                Log.Info("EmployeeJob :: FullAudit", "Check {0} values in EmployeeList.", dataHubEmployee.Name);
                ListItem listItemEmployee = employeeItems.Where(lEmp => lEmp["Title"].ToString() == dataHubEmployee.EmployeeId).First();
                UpdateEmployeeItemValues(listItemEmployee, dataHubEmployee);
            }

            ctx.ExecuteQueryRetry();
        }

        private void ChangeSinceLastRun(ClientContext ctx)
        {
            List employeeList = ctx.Web.GetListByTitle(listName);

            ListItem[] employeeItems = GetEmployeesFromList(ctx, employeeList)
                .Where(eItem => eItem["Status"] == null || (eItem["Status"].ToString() != "Terminated" && eItem["Status"].ToString() != "Left ")).ToArray();
            DataHubResponse employeeIdsResponse = hubRequests.GetEmployeesIdsOnlyAsync(rqParams).Result;

            StringCollection jobCodeParams = new StringCollection() { "department.functionCode=060", $"jobCodeLastUpdated=greaterOrEqual::{previousRun}" };
            DataHubResponse jobCodeUpdateResponse = hubRequests.RunQueryAsync(jobCodeParams).Result;

            StringCollection departmentParams = new StringCollection() { "department.functionCode=060", $"departmentLastUpdated=greaterOrEqual::{previousRun}" };
            DataHubResponse departmentUpdateResponse = hubRequests.RunQueryAsync(departmentParams).Result;

            StringCollection locationParams = new StringCollection() { "department.functionCode=060", $"locationLastUpdated=greaterOrEqual::{previousRun}" };
            DataHubResponse locationUpdateResponse = hubRequests.RunQueryAsync(locationParams).Result;

            StringCollection nameParams = new StringCollection() { "department.functionCode=060", $"nameLastUpdated=greaterOrEqual::{previousRun}" };
            DataHubResponse nameUpdateResponse = hubRequests.RunQueryAsync(nameParams).Result;

            StringCollection phoneParams = new StringCollection() { "department.functionCode=060", $"workPhoneLastUpdated=greaterOrEqual::{previousRun}" };
            DataHubResponse phoneUpdateResponse = hubRequests.RunQueryAsync(phoneParams).Result;

            StringCollection jobDataParams = new StringCollection() { "department.functionCode=060", $"jobDataLastUpdated=greaterOrEqual::{previousRun}" };
            DataHubResponse jobDataUpdateResponse = hubRequests.RunQueryAsync(jobDataParams).Result;

            JoinedOrLeft(employeeIdsResponse, employeeItems, ctx, employeeList);
            
            // Find the 
            Employee[] employeesToUpdate = jobCodeUpdateResponse.Employees
                .Union(departmentUpdateResponse.Employees)
                .Union(locationUpdateResponse.Employees)
                .Union(nameUpdateResponse.Employees)
                .Union(phoneUpdateResponse.Employees)
                .Union(jobDataUpdateResponse.Employees)
                .ToArray();

            foreach (Employee dataHubEmployee in employeesToUpdate)
            {
                Log.Info("EmployeeJob :: SinceLastRun", "Check {0} values in EmployeeList.", dataHubEmployee.Name);
                ListItem listItemEmployee = employeeItems.Where(lEmp => lEmp["Title"].ToString() == dataHubEmployee.EmployeeId).First();
                UpdateEmployeeItemValues(listItemEmployee, dataHubEmployee);
            }

            ctx.ExecuteQueryRetry();
        }
        #endregion

        #region Helper Methods
        private DataHubResponse GetDataHubResponse()
        {
            DataHubResponse response = hubRequests.RunQueryAsync(rqParams).Result;
            Log.Info("EmployeeJob", "DataHub User found {0}", response.Employees[0].Name);
            return response;
        }

        private ListItem[] GetEmployeesFromList(ClientContext ctx, List list)
        {
            CamlQuery query = CamlQuery.CreateAllItemsQuery();
            ListItemCollection employees = list.GetItems(query);

            ctx.Load(employees);
            ctx.ExecuteQueryRetry();
            return employees.ToArray();
        }

        void AddEmployeeToList(List employeeList, Employee employee)
        {
            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
            ListItem newEmployee = employeeList.AddItem(itemCreateInfo);

            foreach (PropertyInfo prop in employeeProps)
            {
                if (employeeList.Fields.GetFieldByInternalName(prop.Name) != null)
                {
                    newEmployee[prop.Name] = employee.GetType().GetProperty(prop.Name).GetValue(employee);
                }
            }

            foreach (PropertyInfo prop in departmentProps)
            {
                string spProp;

                if (prop.Name == "Id" || prop.Name == "Name" || prop.Name == "Code")
                {
                    spProp = $"Department{prop.Name}";
                }
                else
                {
                    spProp = prop.Name;
                }

                if (employeeList.Fields.GetFieldByInternalName(spProp) != null)
                {
                    newEmployee[spProp] = employee.Department.GetType().GetProperty(prop.Name).GetValue(employee.Department);
                }
            }

            newEmployee["Title"] = employee.GetType().GetProperty("EmployeeId").GetValue(employee);

            newEmployee.Update();
        }

        void UpdateEmployeeItemValues(ListItem listItemEmployee, Employee dataHubEmployee)
        {
            bool updateItem = false;

            foreach (PropertyInfo prop in employeeProps)
            {
                if (listItemEmployee.ParentList.Fields.GetFieldByInternalName(prop.Name) != null)
                {
                    string listItemProperty = listItemEmployee[prop.Name] != null ? listItemEmployee[prop.Name].ToString() : String.Empty;
                    object dataHubProperty = dataHubEmployee.GetType().GetProperty(prop.Name).GetValue(dataHubEmployee);
                    string dataHubPropertyString = dataHubProperty!=null ? dataHubProperty.ToString(): String.Empty;

                    if (!updateItem && listItemEmployee[prop.Name] != null && listItemEmployee[prop.Name].GetType() == typeof(DateTime))
                    {
                        updateItem = DateTime.Parse(listItemProperty).ToString() != DateTime.Parse(dataHubPropertyString).ToString();
                    }
                    else if (listItemProperty != dataHubPropertyString)
                    {
                        updateItem = true;
                    }

                    if (updateItem)
                    {
                        listItemEmployee[prop.Name] = dataHubEmployee.GetType().GetProperty(prop.Name).GetValue(dataHubEmployee);
                    }
                }
            }

            foreach (PropertyInfo prop in departmentProps)
            {
                string spProp;

                if (prop.Name == "Id" || prop.Name == "Name" || prop.Name == "Code")
                {
                    spProp = $"Department{prop.Name}";
                }
                else
                {
                    spProp = prop.Name;
                }


                if (listItemEmployee.ParentList.Fields.GetFieldByInternalName(spProp) != null)
                {
                    string listItemProperty = listItemEmployee[spProp] != null ? listItemEmployee[spProp].ToString() : String.Empty;
                    object dataHubProperty = dataHubEmployee.Department.GetType().GetProperty(prop.Name).GetValue(dataHubEmployee.Department);
                    string dataHubPropertyString = dataHubProperty != null ? dataHubProperty.ToString() : String.Empty;

                    if (listItemEmployee[spProp] != null && listItemEmployee[spProp].GetType() == typeof(DateTime))
                    {
                        updateItem = DateTime.Parse(listItemProperty).ToString() != DateTime.Parse(dataHubPropertyString).ToString();
                    }
                    else if (listItemProperty != dataHubPropertyString)
                    {
                        updateItem = true;
                    }

                    if (updateItem)
                    {
                        listItemEmployee[spProp] = dataHubEmployee.Department.GetType().GetProperty(prop.Name).GetValue(dataHubEmployee.Department);
                    }
                }
            }

            if(listItemEmployee["Title"] != null && listItemEmployee["Title"].ToString() != dataHubEmployee.GetType().GetProperty("EmployeeId").GetValue(dataHubEmployee).ToString())
            {
                updateItem = true;
                listItemEmployee["Title"] = dataHubEmployee.GetType().GetProperty("EmployeeId").GetValue(dataHubEmployee);
            }

            if (updateItem) { listItemEmployee.Update(); }
        }
        #endregion
    }
}
