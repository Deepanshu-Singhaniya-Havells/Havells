using System;
using System.Text;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Office.Interop.Excel;
using excel = Microsoft.Office.Interop.Excel;

namespace D365WebJobs
{
    public class SFA_LeadsDataCorrection
    {
        private const string connStr = "AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";
        #region Global Varialble declaration
        static IOrganizationService _service;
        static Guid loginUserGuid = Guid.Empty;
        static string _userId = string.Empty;
        static string _password = string.Empty;
        static string _soapOrganizationServiceUri = string.Empty;
        #endregion

        static void Main(string[] args)
        {
            _service = ConnectToCRM(string.Format(connStr, "https://havells.crm8.dynamics.com"));
            if (((Microsoft.Xrm.Tooling.Connector.CrmServiceClient)_service).IsReady)
            {
                string filePath = @"C:\Kuldeep khare\lead.xlsx";
                string conn = string.Empty;
                Application excelApp = new Application();
                if (excelApp != null)
                {
                    Workbook excelWorkbook = excelApp.Workbooks.Open(filePath, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    Worksheet excelWorksheet = (Worksheet)excelWorkbook.Sheets[1];
                    Range excelRange = excelWorksheet.UsedRange;
                    Range range;

                    string LeadId = string.Empty;
                    for (int i = 1; i <= excelRange.Rows.Count; i++)
                    {
                        try
                        {
                            #region reading Values from Excel file and declaration of local variables 
                            range = (excelWorksheet.Cells[i, 1] as Range);
                            LeadId = range.Value.ToString();
                            LeadId = LeadId.Replace("'", "");
                            #endregion
                            string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='lead'>
                                <attribute name='leadid' />
                                <attribute name='hil_landmark' />
                                <attribute name='createdon' />
                                <attribute name='hil_ticketnumber' />
                                <order attribute='createdon' descending='false' />
                                <filter type='and'>
                                  <condition attribute='hil_ticketnumber' operator='eq' value='{LeadId}' />
                                </filter>
                              </entity>
                            </fetch>";
                            
                            EntityCollection entColl = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                            int j = 0;
                            string _LeadIdRunning = string.Empty;
                            if (entColl.Entities.Count > 2) {
                                Console.WriteLine("3 Duplicate Entires has been found...");
                            }
                            foreach (Entity ent in entColl.Entities) {
                                j += 1;
                                if (j > 1)
                                {
                                    _LeadIdRunning = LeadId + "_" + (j-1).ToString().PadLeft(2, '0');
                                }
                                else {
                                    continue;
                                }
                                ent["hil_ticketnumber"] = _LeadIdRunning;
                                ent["hil_landmark"] = " Duplicate Enquiry Generated from Microsites";
                                _service.Update(ent);

                                //_service.Update(new Entity()
                                //{
                                //    LogicalName = ent.LogicalName,
                                //    Id = ent.Id,
                                //    Attributes = new AttributeCollection() { new System.Collections.Generic.KeyValuePair<string, object>("statecode", new OptionSetValue(1)), new System.Collections.Generic.KeyValuePair<string, object>("statuscode", new OptionSetValue(2)) }
                                //});
                                Console.WriteLine("Processing.. " + i.ToString() + "/" + j.ToString());
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex);
                        }
                        Console.WriteLine("Record Updated.. " + i.ToString() + "|" + LeadId);
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
