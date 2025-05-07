using System;
using Havells.CRM.WebJob.MasterSync.Helper;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Text;


namespace Havells.CRM.WebJob.MasterSync
{
    public class Area
    {
        public static void syncArea(IOrganizationService service)
        {
            try
            {
                Console.WriteLine("Sync Area Function Started ");
                IntegrationConfig intConfig = iHelper.IntegrationConfiguration(service, "Area");
                string authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(intConfig.Auth));
                string timestamp = "19990101000000";
                QueryExpression query = new QueryExpression("hil_area");
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
                                    };

                Dictionary<string, string> parameter = new Dictionary<string, string>();
                var data = iHelper.exesuteAPI(_APIURL, header, parameter, RestSharp.Method.GET);
                var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<AreaRoot>(data.Content);
                int iDone = 0;
                int iCreateCount = 0;
                int iUpdateCount = 0;
                int iTotal = obj.Results.Count;
                if (obj.Success)
                {
                    Console.WriteLine("API Resonse is Sucess total Records " + iTotal);

                    foreach (AreaResult _area in obj.Results)
                    {
                        try
                        {
                            iDone += 1;
                            Entity areaId = iHelper.retriveData("hil_area", "hil_areacode", _area.DM_AREA,service);
                            Entity _entObj = new Entity("hil_area");
                            if (_area.dm_area_desc != null)
                                _entObj["hil_name"] = _area.dm_area_desc;
                            if (_area.DM_AREA != null)
                                _entObj["hil_areacode"] = _area.DM_AREA;
                            if (_area.eff_frmdt != null)
                                _entObj["hil_effectivefromdate"] = Convert.ToDateTime(_area.eff_frmdt);
                            if (_area.eff_todt != null)
                                _entObj["hil_effectivetodate"] = Convert.ToDateTime(_area.eff_todt);
                            if (_area.Mtimestamp == null)
                                _entObj["hil_mdmtimestamp"] = iHelper.StringToDateTime(_area.Ctimestamp);
                            else
                                _entObj["hil_mdmtimestamp"] = iHelper.StringToDateTime(_area.Mtimestamp);

                            if (areaId == null)
                            {
                                if (_area.delete_flag != "X")
                                    service.Create(_entObj);
                                iCreateCount += 1;
                                Console.WriteLine("Area Created For " + _area.DM_AREA + " : " + iCreateCount.ToString() + "/" + iTotal.ToString());
                            }
                            else
                            {
                                if (_area.delete_flag.ToUpper() == "X")
                                { // To Deactivate existing Record in D365
                                    SetStateRequest state = new SetStateRequest();
                                    state.State = new OptionSetValue(1);
                                    state.Status = new OptionSetValue(2);
                                    state.EntityMoniker = new EntityReference("hil_area", areaId.Id);
                                    SetStateResponse stateSet = (SetStateResponse)service.Execute(state);
                                }
                                else
                                {

                                    _entObj.Id = areaId.Id;
                                    service.Update(_entObj);
                                    if (areaId.GetAttributeValue<OptionSetValue>("statecode").Value == 1)
                                    {
                                        SetStateRequest state = new SetStateRequest();
                                        state.State = new OptionSetValue(0);
                                        state.Status = new OptionSetValue(1);
                                        state.EntityMoniker = new EntityReference("hil_area", areaId.Id);
                                        SetStateResponse stateSet = (SetStateResponse)service.Execute(state);
                                    }

                                }
                                iUpdateCount += 1;
                                Console.WriteLine("Area Master Updated For " + _area.DM_AREA + " : " + iUpdateCount.ToString() + "/" + iTotal.ToString());
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
