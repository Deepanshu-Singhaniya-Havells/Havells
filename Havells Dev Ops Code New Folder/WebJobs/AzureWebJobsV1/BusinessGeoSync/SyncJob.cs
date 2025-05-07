using Microsoft.Xrm.Sdk;
using System;
using System.Text;
using Microsoft.Xrm.Sdk.Query;
using System.Collections.Generic;
using Microsoft.Crm.Sdk.Messages;

namespace BusinessGeoSync
{
    public class SyncJob
    {
        public static void synsState(IOrganizationService service)
        {
            try
            {
                Console.WriteLine("Sync State Function Started ");
                IntegrationConfig intConfig = Helper.IntegrationConfiguration(service, "States");
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
                        timestamp = Helper.DateTimeToString(_cTimeStamp);
                    }
                }
                String _APIURL = intConfig.uri + timestamp;
                Dictionary<string, string> header = new Dictionary<string, string>(){
                    { "Authorization", authInfo}
                };

                Dictionary<string, string> parameter = new Dictionary<string, string>();
                var data = Helper.exesuteAPI(_APIURL, header, parameter, RestSharp.Method.GET);
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
                        try
                        {
                            iDone += 1;
                            Entity stateId = Helper.retriveData("hil_state", "hil_statecode", _state.dm_state, service);
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

                            _entObj["hil_country"] = new EntityReference("hil_country", new Guid(Helper._countryId));

                            if (_state.Mtimestamp == null)
                                _entObj["hil_mdmtimestamp"] = Helper.StringToDateTime(_state.Ctimestamp);
                            else
                                _entObj["hil_mdmtimestamp"] = Helper.StringToDateTime(_state.Mtimestamp);

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
                                    Entity _entObjDeleted = new Entity("hil_state", stateId.Id);
                                    if (_state.eff_frmdt != null)
                                        _entObjDeleted["hil_effectivefromdate"] = Convert.ToDateTime(_state.eff_frmdt);
                                    if (_state.eff_todt != null)
                                        _entObjDeleted["hil_effectivetodate"] = Convert.ToDateTime(_state.eff_todt);
                                    if (_state.Mtimestamp == null)
                                        _entObjDeleted["hil_mdmtimestamp"] = Helper.StringToDateTime(_state.Ctimestamp);
                                    else
                                        _entObjDeleted["hil_mdmtimestamp"] = Helper.StringToDateTime(_state.Mtimestamp);
                                    service.Update(_entObjDeleted);

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
        public static void syncDistrict(IOrganizationService service)
        {
            try
            {
                Console.WriteLine("Sync District Function Started ");
                IntegrationConfig intConfig = Helper.IntegrationConfiguration(service, "District");
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
                        timestamp = Helper.DateTimeToString(_cTimeStamp);
                    }
                }
                String _APIURL = intConfig.uri + timestamp;
                string _cityTimeStamp = timestamp;
                Dictionary<string, string> header = new Dictionary<string, string>(){
                    { "Authorization", authInfo}
                };

                Dictionary<string, string> parameter = new Dictionary<string, string>();
                var data = Helper.exesuteAPI(_APIURL, header, parameter, RestSharp.Method.GET);
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
                            Entity districtId = Helper.retriveData("hil_district", "hil_districtcode", _district.dm_dist, service);
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
                            if (_district.Mtimestamp == null)
                                _entObj["hil_mdmtimestamp"] = Helper.StringToDateTime(_district.Ctimestamp);
                            else
                                _entObj["hil_mdmtimestamp"] = Helper.StringToDateTime(_district.Mtimestamp);

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
                                    Entity _entObjDeleted = new Entity("hil_district", districtId.Id);
                                    if (_district.eff_frmdt != null)
                                        _entObjDeleted["hil_effectivefromdate"] = Convert.ToDateTime(_district.eff_frmdt);
                                    if (_district.eff_todt != null)
                                        _entObjDeleted["hil_effectivetodate"] = Convert.ToDateTime(_district.eff_todt);
                                    if (_district.Mtimestamp == null)
                                        _entObjDeleted["hil_mdmtimestamp"] = Helper.StringToDateTime(_district.Ctimestamp);
                                    else
                                        _entObjDeleted["hil_mdmtimestamp"] = Helper.StringToDateTime(_district.Mtimestamp);
                                    service.Update(_entObjDeleted);

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
        public static void syncCity(IOrganizationService service)
        {
            try
            {
                Console.WriteLine("Sync City Function Started ");
                IntegrationConfig intConfig = Helper.IntegrationConfiguration(service, "City");
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
                        timestamp = Helper.DateTimeToString(_cTimeStamp);
                    }
                }
                String _APIURL = intConfig.uri + timestamp;
                Dictionary<string, string> header = new Dictionary<string, string>(){
                    { "Authorization", authInfo}
                };

                Dictionary<string, string> parameter = new Dictionary<string, string>();
                var data = Helper.exesuteAPI(_APIURL, header, parameter, RestSharp.Method.GET);
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
                            Entity cityId = Helper.retriveData("hil_city", "hil_citycode", _city.dm_city, service);
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
                                _entObj["hil_mdmtimestamp"] = Helper.StringToDateTime(_city.Ctimestamp);
                            else
                                _entObj["hil_mdmtimestamp"] = Helper.StringToDateTime(_city.Mtimestamp);

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
                                    Entity _entObjDeleted = new Entity("hil_city", cityId.Id);
                                    if (_city.eff_frmdt != null)
                                        _entObjDeleted["hil_effectivefromdate"] = Convert.ToDateTime(_city.eff_frmdt);
                                    if (_city.eff_todt != null)
                                        _entObjDeleted["hil_effectivetodate"] = Convert.ToDateTime(_city.eff_todt);
                                    if (_city.Mtimestamp == null)
                                        _entObjDeleted["hil_mdmtimestamp"] = Helper.StringToDateTime(_city.Ctimestamp);
                                    else
                                        _entObjDeleted["hil_mdmtimestamp"] = Helper.StringToDateTime(_city.Mtimestamp);
                                    service.Update(_entObjDeleted);

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
        public static void syncPinCode(IOrganizationService service)
        {
            try
            {
                Console.WriteLine("Sync Pincode Function Started ");
                IntegrationConfig intConfig = Helper.IntegrationConfiguration(service, "PinCodes");
                string authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(intConfig.Auth));
                string timestamp = "19990101000000";
                QueryExpression query = new QueryExpression("hil_pincode");
                query.ColumnSet = new ColumnSet("hil_mdmtimestamp");
                query.TopCount = 1;
                query.AddOrder("hil_mdmtimestamp", OrderType.Descending);
                EntityCollection entiColl = service.RetrieveMultiple(query);
                if (entiColl.Entities.Count > 0)
                {
                    if (entiColl[0].Contains("hil_mdmtimestamp"))
                    {
                        DateTime _cTimeStamp = entiColl[0].GetAttributeValue<DateTime>("hil_mdmtimestamp").AddMinutes(330);
                        timestamp = Helper.DateTimeToString(_cTimeStamp);
                    }
                }
                String _APIURL = intConfig.uri + timestamp;
                //String _APIURL = intConfig.uri + "?isinitialload=true";

                string _cityTimeStamp = timestamp;
                Dictionary<string, string> header = new Dictionary<string, string>(){
                    { "Authorization", authInfo}
                };

                Dictionary<string, string> parameter = new Dictionary<string, string>();
                var data = Helper.exesuteAPI(_APIURL, header, parameter, RestSharp.Method.GET);
                var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<PinRoot>(data.Content);
                int iDone = 0;
                int iCreateCount = 0;
                int iUpdateCount = 0;
                int iTotal = obj.Results.Count;
                if (obj.Success)
                {
                    Console.WriteLine("API Resonse is Sucess total Records " + iTotal);

                    foreach (PinResult _pin in obj.Results)
                    {
                        try
                        {
                            iDone += 1;
                            Entity pincodeId = Helper.retriveData("hil_pincode", "hil_name", _pin.dm_pin, service);
                            Entity _entObj = new Entity("hil_pincode");
                            if (_pin.dm_pin != null)
                                _entObj["hil_name"] = _pin.dm_pin;
                            if (_pin.eff_frmdt != null)
                                _entObj["hil_effectivefromdate"] = Convert.ToDateTime(_pin.eff_frmdt);
                            if (_pin.eff_todt != null)
                                _entObj["hil_effectivetodate"] = Convert.ToDateTime(_pin.eff_todt);

                            if (_pin.Mtimestamp == null)
                                _entObj["hil_mdmtimestamp"] = Helper.StringToDateTime(_pin.Ctimestamp);
                            else
                                _entObj["hil_mdmtimestamp"] = Helper.StringToDateTime(_pin.Mtimestamp);

                            if (pincodeId == null)
                            {
                                if (_pin.delete_flag != "X")
                                    service.Create(_entObj);
                                iCreateCount += 1;
                                Console.WriteLine("Pincode Master Created For " + _pin.dm_pin + " : " + iCreateCount.ToString() + "/" + iTotal.ToString());
                            }
                            else
                            {
                                if (_pin.delete_flag.ToUpper() == "X")
                                { // To Deactivate existing Record in D365
                                    Entity _entObjDeleted = new Entity("hil_pincode", pincodeId.Id);
                                    if (_pin.eff_frmdt != null)
                                        _entObjDeleted["hil_effectivefromdate"] = Convert.ToDateTime(_pin.eff_frmdt);
                                    if (_pin.eff_todt != null)
                                        _entObjDeleted["hil_effectivetodate"] = Convert.ToDateTime(_pin.eff_todt);
                                    if (_pin.Mtimestamp == null)
                                        _entObjDeleted["hil_mdmtimestamp"] = Helper.StringToDateTime(_pin.Ctimestamp);
                                    else
                                        _entObjDeleted["hil_mdmtimestamp"] = Helper.StringToDateTime(_pin.Mtimestamp);
                                    service.Update(_entObjDeleted);

                                    SetStateRequest state = new SetStateRequest();
                                    state.State = new OptionSetValue(1);
                                    state.Status = new OptionSetValue(2);
                                    state.EntityMoniker = new EntityReference("hil_pincode", pincodeId.Id);
                                    SetStateResponse stateSet = (SetStateResponse)service.Execute(state);
                                }
                                else
                                {

                                    _entObj.Id = pincodeId.Id;
                                    service.Update(_entObj);
                                    if (pincodeId.GetAttributeValue<OptionSetValue>("statecode").Value == 1)
                                    {
                                        SetStateRequest state = new SetStateRequest();
                                        state.State = new OptionSetValue(0);
                                        state.Status = new OptionSetValue(1);
                                        state.EntityMoniker = new EntityReference("hil_pincode", pincodeId.Id);
                                        SetStateResponse stateSet = (SetStateResponse)service.Execute(state);
                                    }

                                }
                                iUpdateCount += 1;
                                Console.WriteLine("Pin Code Master Updated For " + _pin.dm_pin + " : " + iUpdateCount.ToString() + "/" + iTotal.ToString());
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                            continue;
                        }

                    }
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

        public static void syncArea(IOrganizationService service)
        {
            try
            {
                Console.WriteLine("Sync Area Function Started ");
                IntegrationConfig intConfig = Helper.IntegrationConfiguration(service, "Area");
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
                        timestamp = Helper.DateTimeToString(_cTimeStamp);
                    }
                }
                String _APIURL = intConfig.uri + timestamp;
                //String _APIURL = intConfig.uri + "?isinitialload=true";
                string _cityTimeStamp = timestamp;
                Dictionary<string, string> header = new Dictionary<string, string>(){
                    { "Authorization", authInfo}};

                Dictionary<string, string> parameter = new Dictionary<string, string>();
                var data = Helper.exesuteAPI(_APIURL, header, parameter, RestSharp.Method.GET);
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
                            Entity areaId = Helper.retriveData("hil_area", "hil_areacode", _area.DM_AREA, service);
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
                                _entObj["hil_mdmtimestamp"] = Helper.StringToDateTime(_area.Ctimestamp);
                            else
                                _entObj["hil_mdmtimestamp"] = Helper.StringToDateTime(_area.Mtimestamp);

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
                                    Entity _entObjDeleted = new Entity("hil_area", areaId.Id);
                                    if (_area.eff_frmdt != null)
                                        _entObjDeleted["hil_effectivefromdate"] = Convert.ToDateTime(_area.eff_frmdt);
                                    if (_area.eff_todt != null)
                                        _entObjDeleted["hil_effectivetodate"] = Convert.ToDateTime(_area.eff_todt);
                                    if (_area.Mtimestamp == null)
                                        _entObjDeleted["hil_mdmtimestamp"] = Helper.StringToDateTime(_area.Ctimestamp);
                                    else
                                        _entObjDeleted["hil_mdmtimestamp"] = Helper.StringToDateTime(_area.Mtimestamp);
                                    service.Update(_entObjDeleted);

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
    public class IntegrationConfig
    {
        public string uri { get; set; }
        public string Auth { get; set; }
    }
}

