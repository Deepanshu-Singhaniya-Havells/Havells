using Havells.CRM.WebJob.MasterSync.Helper;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Text;

namespace Havells.CRM.WebJob.MasterSync
{
    public static class City
    {
        public static void syncCity(IOrganizationService service)
        {
            try
            {
                Console.WriteLine("Sync City Function Started ");
                IntegrationConfig intConfig = iHelper.IntegrationConfiguration(service, "City");
                string authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(intConfig.Auth));
                string timestamp = "19990101000000";
                QueryExpression query = new QueryExpression("hil_city");
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
                var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<CityRoot>(data.Content);
                int iDone = 0;
                int iCreateCount = 0;
                int iUpdateCount = 0;
                int iTotal = obj.Results.Count;
                if (obj.Success)
                {
                    Console.WriteLine("API Resonse is Sucess total Records " + iTotal);

                    foreach (CityResult _city in obj.Results)
                    {
                        try
                        {
                            iDone += 1;
                            Entity cityId = iHelper.retriveData("hil_city", "hil_citycode", _city.dm_city, service);
                            Entity _entObj = new Entity("hil_city");
                            if (_city.sap_city_desc != null)
                                _entObj["hil_name"] = _city.sap_city_desc;
                            if (_city.dm_city != null)
                                _entObj["hil_citycode"] = _city.dm_city;
                            if (_city.sap_city_desc != null)
                                _entObj["hil_description"] = _city.sap_city_desc;
                            if (_city.eff_frmdt != null)
                                _entObj["hil_effectivefromdate"] = Convert.ToDateTime(_city.eff_frmdt);
                            if (_city.eff_todt != null)
                                _entObj["hil_effectivetodate"] = Convert.ToDateTime(_city.eff_todt);
                            if (_city.Mtimestamp == null)
                                _entObj["hil_mdmtimestamp"] = iHelper.StringToDateTime(_city.Ctimestamp);
                            else
                                _entObj["hil_mdmtimestamp"] = iHelper.StringToDateTime(_city.Mtimestamp);

                            if (cityId == null)
                            {
                                if (_city.delete_flag != "X")
                                    service.Create(_entObj);
                                iCreateCount += 1;
                                Console.WriteLine("City Master Created For " + _city.dm_city + " : " + iCreateCount.ToString() + "/" + iTotal.ToString());
                            }
                            else
                            {
                                if (_city.delete_flag.ToUpper() == "X")
                                { // To Deactivate existing Record in D365
                                    SetStateRequest state = new SetStateRequest();
                                    state.State = new OptionSetValue(1);
                                    state.Status = new OptionSetValue(2);
                                    state.EntityMoniker = new EntityReference("hil_city", cityId.Id);
                                    SetStateResponse stateSet = (SetStateResponse)service.Execute(state);
                                }
                                else
                                {

                                    _entObj.Id = cityId.Id;
                                    service.Update(_entObj);
                                    if (cityId.GetAttributeValue<OptionSetValue>("statecode").Value == 1)
                                    {
                                        SetStateRequest state = new SetStateRequest();
                                        state.State = new OptionSetValue(0);
                                        state.Status = new OptionSetValue(1);
                                        state.EntityMoniker = new EntityReference("hil_city", cityId.Id);
                                        SetStateResponse stateSet = (SetStateResponse)service.Execute(state);
                                    }

                                }
                                iUpdateCount += 1;
                                Console.WriteLine("City Master Updated For " + _city.dm_city + " : " + iUpdateCount.ToString() + "/" + iTotal.ToString());
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
