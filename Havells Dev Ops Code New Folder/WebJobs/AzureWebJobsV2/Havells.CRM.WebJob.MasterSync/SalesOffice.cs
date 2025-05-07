using System;
using Havells.CRM.WebJob.MasterSync.Helper;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System.Collections.Generic;
using System.Text;

namespace Havells.CRM.WebJob.MasterSync
{
    public class SalesOffice
    {
        public static void syncSalesOffice(IOrganizationService service)
        {
            try
            {
                Console.WriteLine("Sync Sales Office Function Started ");
                IntegrationConfig intConfig = iHelper.IntegrationConfiguration(service, "SalesOffice");
                string authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(intConfig.Auth));
                string timestamp = "19990101000000";
                QueryExpression query = new QueryExpression("hil_salesoffice");
                query.ColumnSet = new ColumnSet("hil_mdmtimestamp");
                query.TopCount = 1;
                query.AddOrder("hil_mdmtimestamp", OrderType.Descending);
                EntityCollection entiColl = service.RetrieveMultiple(query);
                if (entiColl.Entities.Count > 0)
                {
                    if (entiColl[0].Contains("hil_mdmtimestamp"))
                    {
                        DateTime _cTimeStamp = entiColl[0].GetAttributeValue<DateTime>("hil_mdmtimestamp").AddMinutes(330);
                        timestamp = iHelper.DateTimeToString(_cTimeStamp);
                    }
                }
                String _APIURL = intConfig.uri;// + timestamp;
                string _cityTimeStamp = timestamp;
                Dictionary<string, string> header = new Dictionary<string, string>(){
                    { "Authorization", authInfo},
                    {"enquiryDate",_cityTimeStamp }
                };

                Dictionary<string, string> parameter = new Dictionary<string, string>();
                var data = iHelper.exesuteAPI(_APIURL, header, parameter, RestSharp.Method.GET);
                var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<SalesOfficeRoot>(data.Content);
                int iDone = 0;
                int iCreateCount = 0;
                int iUpdateCount = 0;
                int iTotal = obj.Results.Count;
                if (obj.Success)
                {
                    Console.WriteLine("API Resonse is Sucess total Records " + iTotal);

                    foreach (SalesOfficeResult _salesoffice in obj.Results)
                    {
                        try
                        {
                            iDone += 1;
                            Entity salesofficeId = iHelper.retriveData("hil_salesoffice", "hil_salesofficecode", _salesoffice.DM_SALES_OFFICE,service);
                            Entity _entObj = new Entity("hil_salesoffice");
                            if (_salesoffice.SAP_SALES_OFFICE_DESC != null)
                                _entObj["hil_name"] = _salesoffice.SAP_SALES_OFFICE_DESC;
                            if (_salesoffice.DM_SALES_OFFICE != null)
                                _entObj["hil_salesofficecode"] = _salesoffice.DM_SALES_OFFICE;
                            if (_salesoffice.SAP_SALES_OFFICE_DESC != null)
                                _entObj["hil_description"] = _salesoffice.SAP_SALES_OFFICE_DESC;
                            if (_salesoffice.EFF_FRMDT != null)
                                _entObj["hil_effectivefromdate"] = Convert.ToDateTime(_salesoffice.EFF_FRMDT);
                            if (_salesoffice.EFF_TODT != null)
                                _entObj["hil_effectivetodate"] = Convert.ToDateTime(_salesoffice.EFF_TODT);
                            if (_salesoffice.MTIMESTAMP == null)
                                _entObj["hil_mdmtimestamp"] = iHelper.StringToDateTime(_salesoffice.CTIMESTAMP);
                            else
                                _entObj["hil_mdmtimestamp"] = iHelper.StringToDateTime(_salesoffice.MTIMESTAMP);

                            if (salesofficeId == null)
                            {
                                if (_salesoffice.DELETE_FLAG != "X")
                                    service.Create(_entObj);
                                iCreateCount += 1;
                                Console.WriteLine("City Master Created For " + _salesoffice.DM_SALES_OFFICE + " : " + iCreateCount.ToString() + "/" + iTotal.ToString());
                            }
                            else
                            {
                                if (_salesoffice.DELETE_FLAG.ToUpper() == "X")
                                { // To Deactivate existing Record in D365
                                    SetStateRequest state = new SetStateRequest();
                                    state.State = new OptionSetValue(1);
                                    state.Status = new OptionSetValue(2);
                                    state.EntityMoniker = new EntityReference("hil_salesoffice", salesofficeId.Id);
                                    SetStateResponse stateSet = (SetStateResponse)service.Execute(state);
                                }
                                else
                                {

                                    _entObj.Id = salesofficeId.Id;
                                    service.Update(_entObj);
                                    if (salesofficeId.GetAttributeValue<OptionSetValue>("statecode").Value == 1)
                                    {
                                        SetStateRequest state = new SetStateRequest();
                                        state.State = new OptionSetValue(0);
                                        state.Status = new OptionSetValue(1);
                                        state.EntityMoniker = new EntityReference("hil_salesoffice", salesofficeId.Id);
                                        SetStateResponse stateSet = (SetStateResponse)service.Execute(state);
                                    }

                                }
                                iUpdateCount += 1;
                                Console.WriteLine("Sales Office Master Updated For " + _salesoffice.DM_SALES_OFFICE + " : " + iUpdateCount.ToString() + "/" + iTotal.ToString());
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                            continue;
                        }

                    }
                    //Entity state = new Entity("hil_state");

                }
                else
                {
                    Console.WriteLine("Error Code " + obj.ErrorCode);
                    Console.WriteLine("Error Message " + obj.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("SyncState Error : " + ex.Message);
            }
        }

    }
}
