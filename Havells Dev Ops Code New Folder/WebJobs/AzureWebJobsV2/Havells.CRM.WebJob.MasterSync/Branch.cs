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
   public class Branch
    {
      public  static void syncBranch(IOrganizationService service)
        {
            try
            {
                Console.WriteLine("Sync Branch Function Started ");
                IntegrationConfig intConfig = iHelper.IntegrationConfiguration(service, "Branch");
                string authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(intConfig.Auth));
                string timestamp = "19990101000000";
                QueryExpression query = new QueryExpression("hil_branch");
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
                var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<BranchRoot>(data.Content);
                int iDone = 0;
                int iCreateCount = 0;
                int iUpdateCount = 0;
                int iTotal = obj.Results.Count;
                if (obj.Success)
                {
                    Console.WriteLine("API Resonse is Sucess total Records " + iTotal);

                    foreach (BranchResult _branch in obj.Results)
                    {
                        if (_branch.dm_branch == "GJ1")
                        {
                            try
                            {
                                iDone += 1;
                                Entity branchId =iHelper.retriveData("hil_branch", "hil_branchcode", _branch.dm_branch,service);
                                Entity _entObj = new Entity("hil_branch");
                                if (_branch.dm_branch_desc != null)
                                    _entObj["hil_name"] = _branch.dm_branch_desc;
                                if (_branch.dm_branch != null)
                                    _entObj["hil_branchcode"] = _branch.dm_branch;
                                if (_branch.eff_frmdt != null)
                                    _entObj["hil_effectivefromdate"] = Convert.ToDateTime(_branch.eff_frmdt);
                                if (_branch.eff_todt != null)
                                    _entObj["hil_effectivetodate"] = Convert.ToDateTime(_branch.eff_todt);
                                if (_branch.Mtimestamp == null)
                                    _entObj["hil_mdmtimestamp"] = iHelper.StringToDateTime(_branch.Ctimestamp);
                                else
                                    _entObj["hil_mdmtimestamp"] = iHelper.StringToDateTime(_branch.Mtimestamp);

                                if (branchId == null)
                                {
                                    if (_branch.delete_flag != "X")
                                        service.Create(_entObj);
                                    iCreateCount += 1;
                                    Console.WriteLine("Branch Created For " + _branch.dm_branch + " : " + iCreateCount.ToString() + "/" + iTotal.ToString());
                                }
                                else
                                {
                                    if (_branch.delete_flag.ToUpper() == "X")
                                    { // To Deactivate existing Record in D365
                                        SetStateRequest state = new SetStateRequest();
                                        state.State = new OptionSetValue(1);
                                        state.Status = new OptionSetValue(2);
                                        state.EntityMoniker = new EntityReference("hil_branch", branchId.Id);
                                        SetStateResponse stateSet = (SetStateResponse)service.Execute(state);
                                    }
                                    else
                                    {

                                        _entObj.Id = branchId.Id;
                                        service.Update(_entObj);
                                        if (branchId.GetAttributeValue<OptionSetValue>("statecode").Value == 1)
                                        {
                                            SetStateRequest state = new SetStateRequest();
                                            state.State = new OptionSetValue(0);
                                            state.Status = new OptionSetValue(1);
                                            state.EntityMoniker = new EntityReference("hil_branch", branchId.Id);
                                            SetStateResponse stateSet = (SetStateResponse)service.Execute(state);
                                        }

                                    }
                                    iUpdateCount += 1;
                                    Console.WriteLine("Branch Master Updated For " + _branch.dm_branch + " : " + iUpdateCount.ToString() + "/" + iTotal.ToString());
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
