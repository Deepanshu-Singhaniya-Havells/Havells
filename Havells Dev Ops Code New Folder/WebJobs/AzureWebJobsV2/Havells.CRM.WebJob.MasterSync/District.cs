using Havells.CRM.WebJob.MasterSync.Helper;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Text;

namespace Havells.CRM.WebJob.MasterSync
{
    public class District
    {
        public static void syncDistrict(IOrganizationService service)
        {
            try
            {
                Console.WriteLine("Sync District Function Started ");
                IntegrationConfig intConfig = iHelper.IntegrationConfiguration(service, "District");
                string authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(intConfig.Auth));
                string timestamp = "19990101000000";
                QueryExpression query = new QueryExpression("hil_district");
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
                var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<DistrictRoot>(data.Content);
                int iDone = 0;
                int iCreateCount = 0;
                int iUpdateCount = 0;
                int iTotal = obj.Results.Count;
                if (obj.Success)
                {
                    Console.WriteLine("API Resonse is Sucess total Records " + iTotal);

                    foreach (DistrictResult _district in obj.Results)
                    {
                        try
                        {
                            iDone += 1;
                            Entity districtId = iHelper.retriveData("hil_district", "hil_districtcode", _district.dm_dist, service);
                            Entity _entObj = new Entity("hil_district");
                            if (_district.sap_dist_desc != null)
                                _entObj["hil_name"] = _district.sap_dist_desc;
                            if (_district.dm_dist != null)
                                _entObj["hil_districtcode"] = _district.dm_dist;
                            if (_district.sap_dist_desc != null)
                                _entObj["hil_description"] = _district.sap_dist_desc;
                            if (_district.eff_frmdt != null)
                                _entObj["hil_effectivefromdate"] = Convert.ToDateTime(_district.eff_frmdt);
                            if (_district.eff_todt != null)
                                _entObj["hil_effectivetodate"] = Convert.ToDateTime(_district.eff_todt);
                            //  _entObj["hil_country"] = new EntityReference("hil_country", new Guid("F702AE42-E893-E911-A957-000D3AF06C56"));
                            if (_district.Mtimestamp == null)
                                _entObj["hil_mdmtimestamp"] = iHelper.StringToDateTime(_district.Ctimestamp);
                            else
                                _entObj["hil_mdmtimestamp"] = iHelper.StringToDateTime(_district.Mtimestamp);

                            if (districtId == null)
                            {
                                if (_district.delete_flag != "X")
                                    service.Create(_entObj);
                                iCreateCount += 1;
                                Console.WriteLine("District Master Created For " + _district.dm_dist + " : " + iCreateCount.ToString() + "/" + iTotal.ToString());
                            }
                            else
                            {
                                if (_district.delete_flag.ToUpper() == "X")
                                { // To Deactivate existing Record in D365
                                    SetStateRequest state = new SetStateRequest();
                                    state.State = new OptionSetValue(1);
                                    state.Status = new OptionSetValue(2);
                                    state.EntityMoniker = new EntityReference("hil_district", districtId.Id);
                                    SetStateResponse stateSet = (SetStateResponse)service.Execute(state);
                                }
                                else
                                {

                                    _entObj.Id = districtId.Id;
                                    service.Update(_entObj);
                                    if (districtId.GetAttributeValue<OptionSetValue>("statecode").Value == 1)
                                    {
                                        SetStateRequest state = new SetStateRequest();
                                        state.State = new OptionSetValue(0);
                                        state.Status = new OptionSetValue(1);
                                        state.EntityMoniker = new EntityReference("hil_district", districtId.Id);
                                        SetStateResponse stateSet = (SetStateResponse)service.Execute(state);
                                    }

                                }
                                iUpdateCount += 1;
                                Console.WriteLine("District Master Updated For " + _district.dm_dist + " : " + iUpdateCount.ToString() + "/" + iTotal.ToString());
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
