using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using System;
using System.Configuration;
using System.Net;
using System.Text;
using System.ServiceModel.Description;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using Excel = Microsoft.Office.Interop.Excel;

namespace ClaimGeneration
{
    public class Program
    {
        private const string connStr = "AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";

        #region Global Varialble declaration
        static IOrganizationService _service;
        static Guid loginUserGuid = Guid.Empty;
        static string _userId = string.Empty;
        static string _password = string.Empty;
        static string _soapOrganizationServiceUri = string.Empty;
        static string _salesOfficeId = string.Empty;
        #endregion

        static void Main(string[] args)
        {
            _service = ConnectToCRM(string.Format(connStr, "https://havells.crm8.dynamics.com"));
            if (((Microsoft.Xrm.Tooling.Connector.CrmServiceClient)_service).IsReady)
            {
                /* STEPS TO CALCULATE FRANCHISE CLAIM
                Step 1: Rerun Claim Parameter Refresh Batch for pending(Delta) Jobs
                Step 2: Rerun Claim Line Generation Batch for pending(Delta) Jobs
                Step 3: Run Upcountry Classification Flag Batch for pending(Delta) Jobs
                Step 4: Rerun Upcountry Classification Flag Batch for pending(Delta) Jobs
                Step 5: Run Refresh Claim Parameters For KKG Audit Failed Jobs
                Step 6: GenerateTATIncentivesResetChannelPartnerFlag
                Step 7: GenerateTATIncentivesClaimASPWise
                 */

                //Step 0: Pre Activity Batches
                ClaimGeneration.Common.GenerateClaimLinesCountryClassification(_service);
                ClaimGeneration.Common.RefreshClaimParametersForKKGAuditFailedJobs(_service);
                //Post Activity BatchesOn 20th Nigth Batches
                //Step 1:
                //ClaimGeneration.Common.RefreshClaimParameters(_service);
                //Step 2:
                //ClaimGeneration.Common.GenerateClaimLines(_service);
                //Step 2.1:
                //ClaimGeneration.Common.RefreshJobClaimStatusBasedOnSAWActivity(_service);
                //Step 3:
                //ClaimGeneration.Common.RejectPendingSAWApprovals(_service);
                //Step 4:
                //ClaimGeneration.Common.RefreshClaimParametersForKKGAuditFailedJobs(_service);
                //Step 5:
                //ClaimGeneration.Common.GenerateClaimLines(_service);
                //Step 6:
                //ClaimGeneration.Common.GenerateTATIncentivesResetChannelPartnerFlag(_service);
                //Step 6.5:
                //ClaimGeneration.Common.DeleteDuplicateClaimLines(_service);
                //Step 7:
                ClaimGeneration.Common.UpdatePerformaInvoice(_service, Guid.Empty);
                //Step 8:
                // Check Job-View->Pending Claim TAT Achievement Slab missing, is record exist other than LLOYD AC Installation then run ClaimAllSO WebJob Batch
                //ClaimGeneration.Common.GenerateTATIncentivesClaimASPWise(_service);

                /* D365 Forms/Views Changes
                    Change view to be changed: 
                    Performa Invoice - Active Claim Header, 
                    Claim Line - Active Claim Lines, 
                    Jobs - Claim Rejected & Claim Approved, 
                    SAW Activity Approval - Rejected SAW Activity,
                 */

                //CheckError();
                //ClaimGeneration.Common.PostClosureSAWActivityApprovals(_service);

                //Update Activity Code on Claim Lines
                //ClaimGeneration.Common.UpdateActivityOnClaimLines(_service);

                //string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                //      <entity name='msdyn_workorder'>
                //        <attribute name='msdyn_name' />
                //        <attribute name='createdon' />
                //        <attribute name='hil_productsubcategory' />
                //        <attribute name='hil_customerref' />
                //        <attribute name='hil_callsubtype' />
                //        <attribute name='msdyn_workorderid' />
                //        <order attribute='msdyn_name' descending='false' />
                //        <filter type='and'>
                //          <condition attribute='hil_fiscalmonth' operator='eq' uiname='202402' uitype='hil_claimperiod' value='{{16156560-E857-EE11-BE6E-6045BDAD470E}}' />
                //          <condition attribute='hil_warrantysubstatus' operator='eq' value='4' />
                //        </filter>
                //        <link-entity name='hil_claimline' from='hil_jobid' to='msdyn_workorderid' link-type='inner' alias='ab' />
                //      </entity>
                //    </fetch>";
                //EntityCollection entCol = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                //int i = 1;
                //foreach (Entity ent in entCol.Entities) {
                //    ClaimGeneration.Common.UpdateActivityOnClaimLines23Feb2024(_service, ent.Id.ToString());
                //    Console.WriteLine("DONE..." + i++.ToString());
                //}
                //ClaimGeneration.Common.DeleteDuplicateInvoices(_service);
            }
        }
        private static void CheckError() {
            Guid _PIId = new Guid("141c6976-0288-ee11-8178-000d3a3e4cba");
            ClaimOperations clmOps = new ClaimOperations();
            Entity entity = _service.Retrieve("hil_claimheader", _PIId, new ColumnSet(true));

            if (entity.Contains("hil_performastatus"))
            {
                int _performaInvoiceStatus = entity.GetAttributeValue<OptionSetValue>("hil_performastatus").Value;
                if (_performaInvoiceStatus == 2 || _performaInvoiceStatus == 3)
                {
                    string _fetxhXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='hil_claimoverheadline'>
                                <attribute name='hil_claimoverheadlineid' />
                                <attribute name='hil_name' />
                                <attribute name='createdon' />
                                <attribute name='hil_productcategory' />
                                <attribute name='hil_callsubtype' />
                                <order attribute='hil_name' descending='false' />
                                <filter type='and'>
                                  <condition attribute='hil_performainvoice' operator='eq' value='{entity.Id}' />
                                  <filter type='or'>
                                    <condition attribute='hil_productcategory' operator='null' />
                                    <condition attribute='hil_callsubtype' operator='null' />
                                  </filter>
                                  <condition attribute='statecode' operator='eq' value='0' />
                                </filter>
                              </entity>
                            </fetch>";

                    EntityCollection entcoll = _service.RetrieveMultiple(new FetchExpression(_fetxhXML));
                    if (entcoll.Entities.Count > 0)
                    {
                        throw new InvalidPluginExecutionException("There are some claim overheads which do not have Product Category/Call Subtype.");
                    }
                }
                if (_performaInvoiceStatus == 2) // Submit for approval
                {
                    clmOps.GenerateFixedCompensationLines(_service, entity.Id);
                    clmOps.UpdatePerformaInvoice(_service, entity.Id);
                }
                else if (_performaInvoiceStatus == 3) // Approve
                {
                    string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                <entity name='hil_claimheader'>
                                <attribute name='hil_claimheaderid' />
                                <filter type='and'>
                                    <condition attribute='hil_claimheaderid' operator='eq' value='{entity.Id}' />
                                </filter>
                                <link-entity name='account' from='accountid' to='hil_franchisee' link-type='inner' alias='cp'>
                                    <attribute name='hil_salesoffice' />
                                    <attribute name='hil_vendorcode' />
                                </link-entity>
                                </entity>
                                </fetch>";
                    EntityCollection entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    if (!entcoll.Entities[0].Contains("cp.hil_salesoffice") || !entcoll.Entities[0].Contains("cp.hil_vendorcode"))
                    {
                        throw new InvalidPluginExecutionException("Sales Office|Vendor Code is missing at Channel Partner Master record. Please Contact to Service CRM Admin.");
                    }

                    QueryExpression queryExp = new QueryExpression("hil_claimline");
                    queryExp.ColumnSet = new ColumnSet(false);
                    queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                    queryExp.Criteria.AddCondition("hil_claimheader", ConditionOperator.Equal, entity.Id);
                    queryExp.Criteria.AddCondition("hil_activitycode", ConditionOperator.Null);
                    queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);

                    entcoll = _service.RetrieveMultiple(queryExp);
                    if (entcoll.Entities.Count > 0)
                    {
                        throw new InvalidPluginExecutionException("There are some claim lines which do not have Activity Code mapped with. Please Contact to Service CRM Admin.");
                    }
                    else
                    {
                        clmOps.GenerateClaimOverHeads(_service, entity.Id);
                        clmOps.UpdatePerformaInvoice(_service, entity.Id);
                        clmOps.GenerateClaimSummary(_service, entity.Id);
                    }
                }

                //clmOps.UpdatePerformaInvoice(service, entity.Id);
            }
        }
        private static void UpdatePerformaInvoice()
        {
            Excel.Application excelApp = new Excel.Application();
            if (excelApp != null)
            {
                //string appDirectory = AppDomain.CurrentDomain.BaseDirectory + "BizGeoData.xlsx";
                string appDirectory = @"C:\Kuldeep khare\ClaimHeader.xlsx";
                Console.WriteLine(appDirectory);
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(appDirectory, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets[1];
                Excel.Range excelRange = excelWorksheet.UsedRange;
                int rowCount = excelRange.Rows.Count;
                int iLeft = excelRange.Rows.Count;
                int colCount = excelRange.Columns.Count;
                int iDone = 0;
                string _JobCode = string.Empty;
                for (int i = 2; i <= rowCount; i++)  //rowCount
                {
                    _JobCode = (excelWorksheet.Cells[i, 1] as Excel.Range).Value.ToString();

                    ClaimGeneration.Common.UpdatePerformaInvoice(_service, new Guid(_JobCode));
                    iLeft = iLeft - 1;
                    Console.WriteLine("Total row affected " + iDone);
                }
                excelWorkbook.Close();
                excelApp.Quit();
            }
        }
        static void UpdateMissingfieldsInWO()
        {
            QueryExpression queryExp;
            EntityCollection entcoll;
            string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
            <entity name='msdyn_workorder'>
            <attribute name='msdyn_name' />
            <attribute name='createdon' />
            <attribute name='hil_productsubcategory' />
            <attribute name='hil_customerref' />
            <attribute name='hil_callsubtype' />
            <attribute name='msdyn_workorderid' />
            <attribute name='msdyn_customerasset' />
            <order attribute='msdyn_name' descending='false' />
            <filter type='and'>
            <condition attribute='msdyn_timeclosed' operator='on-or-after' value='2020-09-21' />
            <condition attribute='msdyn_substatus' operator='eq' value='{1727FA6C-FA0F-E911-A94E-000D3AF060A1}' />
            <condition attribute='hil_isocr' operator='ne' value='1' />
            <condition attribute='hil_brand' operator='ne' value='3' />
            <condition attribute='msdyn_timeclosed' operator='on-or-before' value='2020-10-20' />
            <filter type='or'>
            <condition attribute='hil_callsubtype' operator='null' />
            <condition attribute='hil_productsubcategory' operator='null' />
            <condition attribute='msdyn_customerasset' operator='null' />
            </filter>
            <condition attribute='hil_claimstatus' operator='ne' value='3' />
            </filter>
            <link-entity name='msdyn_workorderincident' from='msdyn_workorder' to='msdyn_workorderid' link-type='inner' alias='aq'>
            <attribute name='msdyn_customerasset' />
            <link-entity name='msdyn_customerasset' from='msdyn_customerassetid' to='msdyn_customerasset' link-type='inner' alias='ar'>
            <attribute name='hil_invoicedate' />
            </link-entity>
            </link-entity>
            </entity>
            </fetch>";

            while (true)
            {
                entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (entcoll.Entities.Count > 0)
                {
                    EntityReference erCustomerAsset = null;
                    DateTime PurchaseDate;
                    Entity entJobUpdate;
                    int i = 0;
                    foreach (Entity entJob in entcoll.Entities)
                    {
                        try
                        {
                            erCustomerAsset = null;
                            PurchaseDate = new DateTime(1900, 1, 1);
                            if (entJob.Attributes.Contains("aq.msdyn_customerasset"))
                            {
                                erCustomerAsset = (EntityReference)entJob.GetAttributeValue<AliasedValue>("aq.msdyn_customerasset").Value;
                            }
                            if (entJob.Attributes.Contains("ar.hil_invoicedate"))
                            {
                                PurchaseDate = ((DateTime)entJob.GetAttributeValue<AliasedValue>("ar.hil_invoicedate").Value).AddMinutes(330);
                            }
                            entJobUpdate = new Entity(msdyn_workorder.EntityLogicalName, entJob.Id);
                            entJobUpdate["msdyn_customerasset"] = erCustomerAsset;
                            entJobUpdate["hil_purchasedate"] = PurchaseDate;
                            _service.Update(entJobUpdate);
                            Console.WriteLine(i++.ToString());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                    }
                }
            }
        }
        #region App Setting Load/CRM Connection
        static IOrganizationService ConnectToCRM(string connectionString)
        {
            IOrganizationService service = null;
            try
            {
                service = new CrmServiceClient(connectionString);
                if (((CrmServiceClient)service).LastCrmException != null && (((CrmServiceClient)service).LastCrmException.Message == "OrganizationWebProxyClient is null" ||
                    ((CrmServiceClient)service).LastCrmException.Message == "Unable to Login to Dynamics CRM"))
                {
                    Console.WriteLine(((CrmServiceClient)service).LastCrmException.Message);
                    throw new Exception(((CrmServiceClient)service).LastCrmException.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while Creating Conn: " + ex.Message);
            }
            return service;

        }
        #endregion
    }
}
