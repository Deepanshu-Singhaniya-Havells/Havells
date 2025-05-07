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
    public class BranchMapping
    {
        public static void syncBranchMapping(IOrganizationService service)
        {
            try
            {
                Console.WriteLine("Sync Branch Function Started ");
                IntegrationConfig intConfig = iHelper.IntegrationConfiguration(service, "BusinessMapping");
                string authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(intConfig.Auth));
                string timestamp = "19990101000000";
                QueryExpression query = new QueryExpression("hil_businessmapping");
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
                    { "Authorization", authInfo}
                };

                Dictionary<string, string> parameter = new Dictionary<string, string>();
                var data = iHelper.exesuteAPI(_APIURL, header, parameter, RestSharp.Method.GET);
                var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<BusinessMappingRoot>(data.Content);
                int iDone = 0;
                int iCreateCount = 0;
                int iUpdateCount = 0;
                int iTotal = obj.Results.Count;
                if (obj.Success)
                {
                    Console.WriteLine("API Resonse is Sucess total Records " + iTotal);

                    foreach (BusinessMappingResult _business in obj.Results)
                    {
                        try
                        {
                            iDone += 1;
                            if (_business.dm_pin != null && _business.DM_AREA != null && _business.dm_city != null && _business.dm_dist != null && _business.dm_state != null && _business.DM_SALES_OFFICE != null)
                            {
                                Entity pincode = iHelper.retriveData("hil_pincode", "hil_name", _business.dm_pin, service);
                                Entity area = iHelper.retriveData("hil_area", "hil_areacode", _business.DM_AREA, service);
                                Entity city = iHelper.retriveData("hil_city", "hil_citycode", _business.dm_city, service);
                                Entity district = iHelper.retriveData("hil_district", "hil_districtcode", _business.dm_dist, service);
                                Entity sstate = iHelper.retriveData("hil_state", "hil_statecode", _business.dm_state, service);
                                Entity salesoffice = iHelper.retriveData("hil_salesoffice", "hil_salesofficecode", _business.DM_SALES_OFFICE, service);
                                Entity region = iHelper.retriveData("hil_region", "hil_regioncode", _business.dm_region, service);
                                Entity branch = iHelper.retriveData("hil_branch", "hil_branchcode", _business.DM_BRANCH, service);
                                Entity subterritory = iHelper.retriveData("hil_subterritory", "hil_subterritorycode", _business.dm_sub_ter, service);

                                if (area != null && city != null && district != null && sstate != null && salesoffice != null && pincode != null)
                                {
                                    QueryExpression busquery = new QueryExpression("hil_businessmapping");
                                    busquery.ColumnSet = new ColumnSet(true);
                                    busquery.Criteria.AddCondition("hil_pincode", ConditionOperator.Equal, pincode.Id);
                                    busquery.Criteria.AddCondition("hil_area", ConditionOperator.Equal, area.Id);
                                    busquery.Criteria.AddCondition("hil_city", ConditionOperator.Equal, city.Id);
                                    busquery.Criteria.AddCondition("hil_district", ConditionOperator.Equal, district.Id);
                                    busquery.Criteria.AddCondition("hil_state", ConditionOperator.Equal, sstate.Id);
                                    busquery.Criteria.AddCondition("hil_salesoffice", ConditionOperator.Equal, salesoffice.Id);
                                    EntityCollection entColBusiness = service.RetrieveMultiple(busquery);
                                    Entity businessId = null;
                                    if (entColBusiness.Entities.Count > 0)
                                    {
                                        businessId = entColBusiness.Entities[0];
                                    }
                                    string hilname = _business.dm_pin + " " + _business.dm_area_desc + " " + _business.sap_city_desc + " " + _business.sap_dist_desc + " " + _business.sap_state_desc + " " + _business.SAP_SALES_OFFICE_DESC;
                                    Entity _entObj = new Entity("hil_businessmapping");
                                    if (_business.dm_pin != null && _business.DM_AREA != null && _business.dm_city != null && _business.dm_dist != null && _business.dm_state != null && _business.DM_SALES_OFFICE != null)
                                        _entObj["hil_name"] = hilname;
                                    if (_business.dm_pin != null)
                                        _entObj["hil_pincode"] = new EntityReference("hil_pincode", pincode.Id);
                                    if (_business.DM_AREA != null)
                                        _entObj["hil_area"] = new EntityReference("hil_area", area.Id);
                                    if (_business.dm_city != null)
                                        _entObj["hil_city"] = new EntityReference("hil_city", city.Id);
                                    if (_business.dm_dist != null)
                                        _entObj["hil_district"] = new EntityReference("hil_district", district.Id);
                                    if (_business.dm_state != null)
                                        _entObj["hil_state"] = new EntityReference("hil_state", sstate.Id);
                                    if (_business.DM_SALES_OFFICE != null)
                                        _entObj["hil_salesoffice"] = new EntityReference("hil_salesoffice", salesoffice.Id);
                                    if (_business.dm_region != null)
                                        _entObj["hil_region"] = new EntityReference("hil_region", region.Id);
                                    if (_business.DM_BRANCH != null)
                                        _entObj["hil_branch"] = new EntityReference("hil_branch", branch.Id);
                                    if (_business.dm_sub_ter != null)
                                        _entObj["hil_subterritory"] = new EntityReference("hil_subterritory", subterritory.Id);
                                    if (_business.eff_frmdt != null)
                                        _entObj["hil_effectivefromdate"] = Convert.ToDateTime(_business.eff_frmdt);
                                    if (_business.eff_todt != null)
                                        _entObj["hil_effectivetodate"] = Convert.ToDateTime(_business.eff_todt);
                                    if (_business.Mtimestamp == null)
                                        _entObj["hil_mdmtimestamp"] = iHelper.StringToDateTime(_business.Ctimestamp);
                                    else
                                        _entObj["hil_mdmtimestamp"] = iHelper.StringToDateTime(_business.Mtimestamp);

                                    if (businessId.Id == null)
                                    {
                                        if (_business.delete_flag != "X")
                                            service.Create(_entObj);
                                        iCreateCount += 1;
                                        Console.WriteLine("Business Mapping Created For " + hilname + " : " + iCreateCount.ToString() + "/" + iTotal.ToString());
                                    }
                                    else
                                    {
                                        if (_business.delete_flag.ToUpper() == "X")
                                        { // To Deactivate existing Record in D365
                                            SetStateRequest state = new SetStateRequest();
                                            state.State = new OptionSetValue(1);
                                            state.Status = new OptionSetValue(2);
                                            state.EntityMoniker = new EntityReference("hil_businessmapping", businessId.Id);
                                            SetStateResponse stateSet = (SetStateResponse)service.Execute(state);
                                        }
                                        else
                                        {

                                            _entObj.Id = businessId.Id;
                                            service.Update(_entObj);
                                            if (businessId.GetAttributeValue<OptionSetValue>("statecode").Value == 1)
                                            {
                                                SetStateRequest state = new SetStateRequest();
                                                state.State = new OptionSetValue(0);
                                                state.Status = new OptionSetValue(1);
                                                state.EntityMoniker = new EntityReference("hil_businessmapping", businessId.Id);
                                                SetStateResponse stateSet = (SetStateResponse)service.Execute(state);
                                            }

                                        }
                                        iUpdateCount += 1;
                                        Console.WriteLine("Business Mapping Master Updated For " + hilname + " : " + iUpdateCount.ToString() + "/" + iTotal.ToString());
                                    }

                                }
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
