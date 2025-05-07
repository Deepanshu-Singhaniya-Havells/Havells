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
   public class Region
    {
        public static void syncRegion(IOrganizationService service)
        {
            try
            {
                Console.WriteLine("Sync Region Function Started ");
                IntegrationConfig intConfig = iHelper.IntegrationConfiguration(service, "Regions");
                string authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(intConfig.Auth));
                string timestamp = "19990101000000";
                QueryExpression query = new QueryExpression("hil_region");
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
                var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<RegionRoot>(data.Content);
                int iDone = 0;
                int iCreateCount = 0;
                int iUpdateCount = 0;
                int iTotal = obj.Results.Count;
                if (obj.Success)
                {
                    Console.WriteLine("API Resonse is Sucess total Records " + iTotal);

                    foreach (RegionResult _region in obj.Results)
                    {
                        try
                        {
                            iDone += 1;
                            Entity regionId =iHelper.retriveData("hil_region", "hil_regioncode", _region.dm_region,service);
                            Entity _entObj = new Entity("hil_region");
                            if (_region.sap_region_desc != null)
                                _entObj["hil_name"] = _region.sap_region_desc;
                            if (_region.dm_region != null)
                                _entObj["hil_regioncode"] = _region.dm_region;
                            if (_region.eff_frmdt != null)
                                _entObj["hil_effectivefromdate"] = Convert.ToDateTime(_region.eff_frmdt);
                            if (_region.eff_todt != null)
                                _entObj["hil_effectivetodate"] = Convert.ToDateTime(_region.eff_todt);
                            if (_region.Mtimestamp == null)
                                _entObj["hil_mdmtimestamp"] = iHelper.StringToDateTime(_region.Ctimestamp);
                            else
                                _entObj["hil_mdmtimestamp"] = iHelper.StringToDateTime(_region.Mtimestamp);

                            if (regionId == null)
                            {
                                if (_region.delete_flag != "X")
                                    service.Create(_entObj);
                                iCreateCount += 1;
                                Console.WriteLine("Region Created For " + _region.dm_region + " : " + iCreateCount.ToString() + "/" + iTotal.ToString());
                            }
                            else
                            {
                                if (_region.delete_flag.ToUpper() == "X")
                                { // To Deactivate existing Record in D365
                                    SetStateRequest state = new SetStateRequest();
                                    state.State = new OptionSetValue(1);
                                    state.Status = new OptionSetValue(2);
                                    state.EntityMoniker = new EntityReference("hil_region", regionId.Id);
                                    SetStateResponse stateSet = (SetStateResponse)service.Execute(state);
                                }
                                else
                                {

                                    _entObj.Id = regionId.Id;
                                    service.Update(_entObj);
                                    if (regionId.GetAttributeValue<OptionSetValue>("statecode").Value == 1)
                                    {
                                        SetStateRequest state = new SetStateRequest();
                                        state.State = new OptionSetValue(0);
                                        state.Status = new OptionSetValue(1);
                                        state.EntityMoniker = new EntityReference("hil_region", regionId.Id);
                                        SetStateResponse stateSet = (SetStateResponse)service.Execute(state);
                                    }

                                }
                                iUpdateCount += 1;
                                Console.WriteLine("Region Master Updated For " + _region.dm_region + " : " + iUpdateCount.ToString() + "/" + iTotal.ToString());
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
