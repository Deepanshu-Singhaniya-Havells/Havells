using Havells.CRM.WebJob.MasterSync.Helper;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Text;

namespace Havells.CRM.WebJob.MasterSync
{
    public static class State
    {
       public static void synsState(IOrganizationService service)
        {
            try
            {
                Console.WriteLine("Sync State Function Started ");
                IntegrationConfig intConfig = iHelper.IntegrationConfiguration(service, "States");
                string authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(intConfig.Auth));
                string timestamp = "19990101000000";
                QueryExpression query = new QueryExpression("hil_state");
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
                String _APIURL = intConfig.uri + timestamp;
                Dictionary<string, string> header = new Dictionary<string, string>(){
                    { "Authorization", authInfo}
                };

                Dictionary<string, string> parameter = new Dictionary<string, string>();
                var data = iHelper.exesuteAPI(_APIURL, header, parameter, RestSharp.Method.GET);
                var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<StateRoot>(data.Content);
                int iDone = 0;
                int iCreateCount = 0;
                int iUpdateCount = 0;
                int iTotal = obj.Results.Count;
                if (obj.Success)
                {
                    Console.WriteLine("API Resonse is Sucess total Records " + iTotal);

                    foreach (StateResult _state in obj.Results)
                    {
                        //if (_state.dm_state == "GJ")
                        {
                            try
                            {
                                iDone += 1;
                                Entity stateId = iHelper.retriveData("hil_state", "hil_statecode", _state.dm_state, service);
                                Entity _entObj = new Entity("hil_state");
                                if (_state.sap_state_desc != null)
                                    _entObj["hil_name"] = _state.sap_state_desc;
                                if (_state.dm_state != null)
                                    _entObj["hil_statecode"] = _state.dm_state;
                                if (_state.SAP_state != null)
                                    _entObj["hil_sapstatecode"] = _state.SAP_state;
                                if (_state.GSTStateName != null)
                                    _entObj["hil_gststatename"] = _state.GSTStateName;
                                if (_state.GSTStateCode != null)
                                    _entObj["hil_gststatecode"] = _state.GSTStateCode;
                                if (_state.eff_frmdt != null)
                                    _entObj["hil_effectivefromdate"] = Convert.ToDateTime(_state.eff_frmdt);
                                if (_state.eff_todt != null)
                                    _entObj["hil_effectivetodate"] = Convert.ToDateTime(_state.eff_todt);

                                _entObj["hil_country"] = new EntityReference("hil_country", new Guid("F702AE42-E893-E911-A957-000D3AF06C56"));

                                if (_state.Mtimestamp == null)
                                    _entObj["hil_mdmtimestamp"] = iHelper.StringToDateTime(_state.Ctimestamp);
                                else
                                    _entObj["hil_mdmtimestamp"] = iHelper.StringToDateTime(_state.Mtimestamp);

                                if (stateId == null)
                                {
                                    if (_state.delete_flag != "X")
                                        service.Create(_entObj);
                                    iCreateCount += 1;
                                    Console.WriteLine("State Master Created For " + _state.dm_state + " : " + iCreateCount.ToString() + "/" + iTotal.ToString());
                                }
                                else
                                {
                                    if (_state.delete_flag.ToUpper() == "X")
                                    { // To Deactivate existing Record in D365
                                        SetStateRequest state = new SetStateRequest();
                                        state.State = new OptionSetValue(1);
                                        state.Status = new OptionSetValue(2);
                                        state.EntityMoniker = new EntityReference("hil_state", stateId.Id);
                                        SetStateResponse stateSet = (SetStateResponse)service.Execute(state);
                                    }
                                    else
                                    {

                                        _entObj.Id = stateId.Id;
                                        service.Update(_entObj);
                                        if (stateId.GetAttributeValue<OptionSetValue>("statecode").Value == 1)
                                        {
                                            SetStateRequest state = new SetStateRequest();
                                            state.State = new OptionSetValue(0);
                                            state.Status = new OptionSetValue(1);
                                            state.EntityMoniker = new EntityReference("hil_state", stateId.Id);
                                            SetStateResponse stateSet = (SetStateResponse)service.Execute(state);
                                        }

                                    }
                                    iUpdateCount += 1;
                                    Console.WriteLine("State Master Updated For " + _state.dm_state + " : " + iUpdateCount.ToString() + "/" + iTotal.ToString());
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                                continue;
                            }
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
