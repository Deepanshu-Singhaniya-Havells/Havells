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
    public class Pincode
    {
        public static void syncPinCode(IOrganizationService service)
        {
            try
            {
                Console.WriteLine("Sync Pincode Function Started ");
                IntegrationConfig intConfig = iHelper.IntegrationConfiguration(service, "PinCodes");
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
                            //if (_pin.dm_pin == "700016")
                            //{
                            iDone += 1;
                            Entity pincodeId = iHelper.retriveData("hil_pincode", "hil_name", _pin.dm_pin, service);
                            Entity _entObj = new Entity("hil_pincode");
                            if (_pin.dm_pin != null)
                                _entObj["hil_name"] = _pin.dm_pin;
                            if (_pin.eff_frmdt != null)
                                _entObj["hil_effectivefromdate"] = Convert.ToDateTime(_pin.eff_frmdt);
                            if (_pin.eff_todt != null)
                                _entObj["hil_effectivetodate"] = Convert.ToDateTime(_pin.eff_todt);

                            if (_pin.Mtimestamp == null)
                                _entObj["hil_mdmtimestamp"] = iHelper.StringToDateTime(_pin.Ctimestamp);
                            else
                                _entObj["hil_mdmtimestamp"] = iHelper.StringToDateTime(_pin.Mtimestamp);

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
                                    SetStateRequest state = new SetStateRequest();
                                    state.State = new OptionSetValue(1);
                                    state.Status = new OptionSetValue(2);
                                    state.EntityMoniker = new EntityReference("hil_pincodeid", pincodeId.Id);
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
                                        state.EntityMoniker = new EntityReference("hil_pincodeid", pincodeId.Id);
                                        SetStateResponse stateSet = (SetStateResponse)service.Execute(state);
                                    }

                                }
                                iUpdateCount += 1;
                                Console.WriteLine("Pin Code Master Updated For " + _pin.dm_pin + " : " + iUpdateCount.ToString() + "/" + iTotal.ToString());
                            }
                            // }
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
    }
}
