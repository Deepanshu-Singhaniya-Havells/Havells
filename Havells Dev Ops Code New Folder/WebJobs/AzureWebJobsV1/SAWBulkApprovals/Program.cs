using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Runtime.Serialization.Json;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Client;
using System.Configuration;
using System.ServiceModel.Description;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using Newtonsoft.Json;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace SAWBulkApprovals
{
    class Program
    {
        #region Global Varialble declaration
        static IOrganizationService _service;
        static Guid loginUserGuid = Guid.Empty;
        static string _userId = string.Empty;
        static string _password = string.Empty;
        static string _soapOrganizationServiceUri = string.Empty;
        #endregion
        static void Main(string[] args)
        {
            if (DateTime.Now.Year == 2021 || DateTime.Now.Month <= 10)
            {
                LoadAppSettings();
                if (loginUserGuid != Guid.Empty)
                {
                    //Sync_SAWApprovals();
                    RejectPendingSAWApprovals();
                    //UpdateChannelPartnerTaxNumbers();
                }
            }
            else
            {
                Console.WriteLine("D365 Security token has been refreshed from backend. Please contact to System Admin.");
            }
        }

        #region App Setting Load/CRM Connection
        private static void LoadAppSettings()
        {
            try
            {
                _userId = ConfigurationManager.AppSettings["CrmUserId"].ToString();
                _password = ConfigurationManager.AppSettings["CrmUserPassword"].ToString();
                _soapOrganizationServiceUri = ConfigurationManager.AppSettings["CrmSoapOrganizationServiceUri"].ToString();
                //Console.WriteLine("Please Enter Password :");
                //_password = Console.ReadLine();
                ConnectToCRM();
            }
            catch (Exception ex)
            {
                Console.WriteLine("SAWBulkApprovals.Program.Main.LoadAppSettings ::  Error While Loading App Settings:" + ex.Message.ToString());
            }
        }
        private static void ConnectToCRM()
        {
            try
            {
                Console.WriteLine("Connecting......");
                ClientCredentials credentials = new ClientCredentials();
                credentials.UserName.UserName = _userId;
                credentials.UserName.Password = _password;
                Uri serviceUri = new Uri(_soapOrganizationServiceUri);
                System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                ServicePointManager.ServerCertificateValidationCallback = delegate (object s, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };
                OrganizationServiceProxy proxy = new OrganizationServiceProxy(serviceUri, null, credentials, null);
                proxy.EnableProxyTypes();
                _service = (IOrganizationService)proxy;
                loginUserGuid = ((WhoAmIResponse)_service.Execute(new WhoAmIRequest())).UserId;
                if (loginUserGuid != null)
                {
                    Console.WriteLine("Connection has been stablished successfully.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("SAWBulkApprovals.Program.Main.ConnectToCRM :: Error While Creating Connection with MS CRM Organisation:" + ex.Message.ToString());
            }
        }
        #endregion

        private static void Sync_SAWApprovals()
        {
            Excel.Application excelApp = new Excel.Application();
            if (excelApp != null)
            {
                Console.WriteLine("Reading Excel file......");
                string appDirectory = AppDomain.CurrentDomain.BaseDirectory + "SAW_Approvals.xlsx";
                Console.WriteLine(appDirectory);
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(appDirectory, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets[1];
                Excel.Range excelRange = excelWorksheet.UsedRange;
                int rowCount = excelRange.Rows.Count;
                int iLeft = excelRange.Rows.Count;
                int colCount = excelRange.Columns.Count;
                int iDone = 1;
                string _sawapprovalsId = string.Empty;
                string _Status = string.Empty;
                string _Remarks = string.Empty;
                Entity ent = null;
                for (int i = 2; i <= rowCount; i++)  //rowCount
                {
                    _sawapprovalsId = (excelWorksheet.Cells[i, 1] as Excel.Range).Value.ToString();
                    _Status = (excelWorksheet.Cells[i, 2] as Excel.Range).Value.ToString().ToUpper();
                    _Remarks = (excelWorksheet.Cells[i, 3] as Excel.Range).Value.ToString();
                    ent = new Entity("hil_sawactivityapproval", new Guid(_sawapprovalsId));
                    ent["hil_approvalstatus"] = new OptionSetValue(_Status == "APPROVED" ? 3 : 4);
                    ent["hil_approverremarks"] = _Remarks;
                    _service.Update(ent);
                    Console.WriteLine("Total row affected: " + iDone++);
                }
                Console.WriteLine("Batch has been completed.");
                excelWorkbook.Close();
                excelApp.Quit();
            }
        }

        private static void RejectPendingSAWApprovals()
        {
            QueryExpression queryExp;
            EntityCollection entCol;
            queryExp = new QueryExpression("hil_sawactivityapproval");
            queryExp.ColumnSet = new ColumnSet("hil_approver", "hil_approvalstatus");
            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            queryExp.Criteria.AddCondition("createdon", ConditionOperator.OnOrBefore, "2022-06-20");
            queryExp.Criteria.AddCondition("hil_approvalstatus", ConditionOperator.In, new object[] { 1, 2, 5, 6 });
            entCol = _service.RetrieveMultiple(queryExp);
            foreach (Entity ent in entCol.Entities)
            {
                ent["hil_approver"] = new EntityReference("systemuser", loginUserGuid);
                _service.Update(ent);
                ent["hil_approverremarks"] = "Due to non-review by Branch, SAW & Claim has been rejected for this Job. Please check with your Branch Service Head for the same.";
                ent["hil_approvalstatus"] = new OptionSetValue(4);
                _service.Update(ent);
            }
        }

        private static void UpdateChannelPartnerTaxNumbers()
        {
            Excel.Application excelApp = new Excel.Application();
            if (excelApp != null)
            {
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(@"C:\Kuldeep khare\ChannelPartner.xlsx", 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets[1];
                Excel.Range excelRange = excelWorksheet.UsedRange;
                int rowCount = excelRange.Rows.Count;
                int iLeft = excelRange.Rows.Count;

                int colCount = excelRange.Columns.Count;
                Entity channelpartner = null;
                string GstNo = string.Empty;
                string PanNo = string.Empty;

                for (int i = 2; i <= rowCount; i++)
                {
                    Guid recId = new Guid((excelWorksheet.Cells[i, 1] as Excel.Range).Value == null ? "" : (excelWorksheet.Cells[i, 1] as Excel.Range).Value.ToString());
                    GstNo = (excelWorksheet.Cells[i, 2] as Excel.Range).Value == null ? "" : (excelWorksheet.Cells[i, 2] as Excel.Range).Value.ToString();
                    PanNo = (excelWorksheet.Cells[i, 3] as Excel.Range).Value == null ? "" : (excelWorksheet.Cells[i, 3] as Excel.Range).Value.ToString();

                    try
                    {
                        channelpartner = new Entity("account", recId);
                        channelpartner.Attributes["address1_postofficebox"] = GstNo;
                        channelpartner.Attributes["hil_pan"] = PanNo;
                        _service.Update(channelpartner);
                        Console.WriteLine("Rows affected " + (i - 1).ToString());
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
                excelWorkbook.Close();
                Console.ReadLine();
                excelApp.Quit();
            }
        }
    }
}
