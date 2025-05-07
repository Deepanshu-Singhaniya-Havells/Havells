using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Havells_Plugin.Product_Request_Header
{
    public class PostUpdate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            try
            {
                if (context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == hil_productrequestheader.EntityLogicalName.ToLower()
                    && context.MessageName.ToUpper() == "UPDATE")
                {
                    //hil_productrequestheader PductReqHd = (hil_productrequestheader)context.InputParameters["Target"];
                    Entity entity = (Entity)context.InputParameters["Target"];
                    hil_productrequestheader PductReqHd = entity.ToEntity<hil_productrequestheader>();
                    InitiateExecution(service, PductReqHd);
                }
            }
            catch(Exception ex)
            {
                throw new InvalidPluginExecutionException("ProductRequestHeader.PostUpdate.Execute : " +ex.Message);
            }
        }
        #region Initiate Execution
        public static void InitiateExecution(IOrganizationService service, hil_productrequestheader PductReqHd)
        {
            hil_productrequestheader Head = (hil_productrequestheader)service.Retrieve(hil_productrequestheader.EntityLogicalName, PductReqHd.Id, new ColumnSet(true));
            if (Head.hil_PRType != null)
            {
                if (Head.hil_PRType.Value == 910590002)// FOR PRODUCT REPLACEMENT
                {
                    #region Product Replacement Approval
                    int iSink = 0;
                    int iSource = 0;
                    int iApproval = 0;
                    int iNext = 0;
                    Guid iLineId = new Guid();
                    Guid iSubStatus = new Guid();
                    EntityReference iFall = new EntityReference();
                    msdyn_workorder _thisJob = new msdyn_workorder();
                    if (Head.Attributes.Contains("hil_finallevel"))
                    {
                        OptionSetValue final = (OptionSetValue)Head["hil_finallevel"];
                        iSink = final.Value;
                    }
                    if (Head.Attributes.Contains("hil_pendingat"))
                    {
                        OptionSetValue thisLevel = (OptionSetValue)Head["hil_pendingat"];
                        iSource = thisLevel.Value;
                    }
                    if (PductReqHd.hil_level1status != null)
                    {
                        iApproval = PductReqHd.hil_level1status.Value;
                        if (iSource == 1)
                        {
                            if (iApproval == 1)
                            {
                                _thisJob = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, Head.hil_Job.Id, new ColumnSet("hil_regardingfallback"));
                                if (_thisJob.Contains("hil_regardingfallback") && _thisJob.Attributes.Contains("hil_regardingfallback"))
                                {
                                    iFall = _thisJob.GetAttributeValue<EntityReference>("hil_regardingfallback");
                                }
                                else
                                {
                                    throw new InvalidPluginExecutionException("XXXXXXXXXXXXXXXXXXXXXXXXXXXXX - Fall Back Can't be Null in Job - XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX");
                                }
                                if (iSource == iSink)
                                {
                                    iLineId = GetLineforThisHeader(service, PductReqHd.Id);
                                    //iSubStatus = Helper.GetGuidbyName(msdyn_workordersubstatus.EntityLogicalName, "msdyn_name", "Work Initiated", service);
                                    //if (iSubStatus != Guid.Empty)
                                    //{
                                        hil_productrequestheader Header = new hil_productrequestheader();
                                        Header.Id = PductReqHd.Id;
                                        Header.hil_SyncStatus = new OptionSetValue(1);
                                        Header.statuscode = new OptionSetValue(910590003);
                                        Header["hil_pendingat"] = null;
                                        //Header.hil_SyncStatus = new OptionSetValue(1);
                                        Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Work Initiated", _thisJob.Id, service);
                                        //_thisJob.msdyn_SubStatus = new EntityReference(msdyn_workordersubstatus.EntityLogicalName, iSubStatus);
                                        //service.Update(_thisJob);
                                        service.Update(Header);
                                        if (iLineId != Guid.Empty)
                                        {
                                            hil_productrequest Line = new hil_productrequest();
                                            Line.Id = iLineId;// (hil_productrequest)service.Retrieve(hil_productrequest.EntityLogicalName, iLineId, new ColumnSet(false));
                                            Line.statuscode = new OptionSetValue(910590003);
                                            service.Update(Line);
                                        }
                                    //}
                                }
                                else if (iSource < iSink)
                                {
                                    iNext = 2;
                                    hil_productrequestheader Header = new hil_productrequestheader();
                                    Header.Id = PductReqHd.Id;
                                    Header["hil_pendingat"] = new OptionSetValue(iNext);
                                    service.Update(Header);
                                    iLineId = GetLineforThisHeader(service, PductReqHd.Id);
                                    if (iLineId != Guid.Empty)
                                        AssignToNextLevel(service, Head.hil_Division.Id, iNext, iLineId, PductReqHd.Id, iFall);
                                }
                            }
                            else if (iApproval == 2)
                            {
                                iLineId = GetLineforThisHeader(service, PductReqHd.Id);
                                _thisJob = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, Head.hil_Job.Id, new ColumnSet(false));
                                hil_productrequestheader Header = new hil_productrequestheader();
                                Header.Id = PductReqHd.Id;
                                //Header.hil_SyncStatus = new OptionSetValue(1);
                                Header.statuscode = new OptionSetValue(910590004);
                                Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Work Initiated", _thisJob.Id, service);
                                service.Update(Header);
                                if (iLineId != Guid.Empty)
                                {
                                    hil_productrequest Line = new hil_productrequest();
                                    Line.Id = iLineId;// (hil_productrequest)service.Retrieve(hil_productrequest.EntityLogicalName, iLineId, new ColumnSet(false));
                                    Line.statuscode = new OptionSetValue(910590004);
                                    service.Update(Line);
                                }
                            }
                        }
                    }
                    else if (PductReqHd.hil_Level2Status != null)
                    {
                        iApproval = PductReqHd.hil_Level2Status.Value;
                        if (iSource == 2)
                        {
                            if (iApproval == 1)
                            {
                                _thisJob = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, Head.hil_Job.Id, new ColumnSet("hil_regardingfallback"));
                                if (_thisJob.Contains("hil_regardingfallback") && _thisJob.Attributes.Contains("hil_regardingfallback"))
                                {
                                    iFall = _thisJob.GetAttributeValue<EntityReference>("hil_regardingfallback");
                                }
                                else
                                {
                                    throw new InvalidPluginExecutionException("XXXXXXXXXXXXXXXXXXXXXXXXXXXXX - Fall Back Can't be Null in Job - XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX");
                                }
                                if (iSource == iSink)
                                {
                                    iLineId = GetLineforThisHeader(service, PductReqHd.Id);
                                    //iSubStatus = Helper.GetGuidbyName(msdyn_workordersubstatus.EntityLogicalName, "msdyn_name", "Work Initiated", service);
                                    
                                    //if (iSubStatus != Guid.Empty)
                                    //{
                                        hil_productrequestheader Header = new hil_productrequestheader();
                                        Header.Id = PductReqHd.Id;
                                        Header.hil_SyncStatus = new OptionSetValue(1);
                                        Header.statuscode = new OptionSetValue(910590003);
                                        Header.hil_SyncStatus = new OptionSetValue(1);
                                        Header["hil_pendingat"] = null;
                                        Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Work Initiated", _thisJob.Id, service);
                                        service.Update(Header);
                                        if (iLineId != Guid.Empty)
                                        {
                                            hil_productrequest Line = new hil_productrequest();
                                            Line.Id = iLineId;// (hil_productrequest)service.Retrieve(hil_productrequest.EntityLogicalName, iLineId, new ColumnSet(false));
                                            Line.statuscode = new OptionSetValue(910590003);
                                            service.Update(Line);
                                        }
                                    //}
                                }
                                else if (iSource < iSink)
                                {
                                    iNext = 3;
                                    hil_productrequestheader Header = new hil_productrequestheader();
                                    Header.Id = PductReqHd.Id;
                                    Header["hil_pendingat"] = new OptionSetValue(iNext);
                                    service.Update(Header);
                                    iLineId = GetLineforThisHeader(service, PductReqHd.Id);
                                    if (iLineId != Guid.Empty)
                                        AssignToNextLevel(service, Head.hil_Division.Id, iNext, iLineId, PductReqHd.Id, iFall);
                                }
                            }
                            else if (iApproval == 2)
                            {
                                iLineId = GetLineforThisHeader(service, PductReqHd.Id);
                                _thisJob = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, Head.hil_Job.Id, new ColumnSet(false));
                                    hil_productrequestheader Header = new hil_productrequestheader();
                                    Header.Id = PductReqHd.Id;
                                    //Header.hil_SyncStatus = new OptionSetValue(1);
                                    Header.statuscode = new OptionSetValue(910590004);
                                    //Header.hil_SyncStatus = new OptionSetValue(1);
                                    Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Work Initiated", _thisJob.Id, service);
                                    service.Update(Header);
                                    if (iLineId != Guid.Empty)
                                    {
                                        hil_productrequest Line = new hil_productrequest();
                                        Line.Id = iLineId;// (hil_productrequest)service.Retrieve(hil_productrequest.EntityLogicalName, iLineId, new ColumnSet(false));
                                        Line.statuscode = new OptionSetValue(910590004);
                                        service.Update(Line);
                                    }
                            }
                        }
                    }
                    else if (PductReqHd.hil_Level3Status != null)
                    {
                        iApproval = PductReqHd.hil_Level3Status.Value;
                        if (iSource == 3)
                        {
                            if (iApproval == 1)
                            {
                                if (iSource == iSink)
                                {
                                    iLineId = GetLineforThisHeader(service, PductReqHd.Id);
                                    _thisJob = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, Head.hil_Job.Id, new ColumnSet(false));
                                    hil_productrequestheader Header = new hil_productrequestheader();
                                    Header.Id = PductReqHd.Id;
                                    Header["hil_pendingat"] = null;
                                    //Header.hil_SyncStatus = new OptionSetValue(1);
                                    Header.statuscode = new OptionSetValue(910590003);
                                    Header.hil_SyncStatus = new OptionSetValue(1);
                                    service.Update(Header);
                                    Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Work Initiated", _thisJob.Id, service);
                                    if (iLineId != Guid.Empty)
                                    {
                                        hil_productrequest Line = new hil_productrequest();
                                        Line.Id = iLineId;// (hil_productrequest)service.Retrieve(hil_productrequest.EntityLogicalName, iLineId, new ColumnSet(false));
                                        Line.statuscode = new OptionSetValue(910590003);
                                        service.Update(Line);
                                    }
                                }
                            }
                            else if (iApproval == 2)
                            {
                                iLineId = GetLineforThisHeader(service, PductReqHd.Id);
                                _thisJob = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, Head.hil_Job.Id, new ColumnSet(false));
                                    hil_productrequestheader Header = new hil_productrequestheader();
                                    Header.Id = PductReqHd.Id;
                                    //Header.hil_SyncStatus = new OptionSetValue(1);
                                    Header.statuscode = new OptionSetValue(910590004);
                                    Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Work Initiated", _thisJob.Id, service);
                                    service.Update(Header);
                                    if (iLineId != Guid.Empty)
                                    {
                                        hil_productrequest Line = new hil_productrequest();
                                        Line.Id = iLineId;// (hil_productrequest)service.Retrieve(hil_productrequest.EntityLogicalName, iLineId, new ColumnSet(false));
                                        Line.statuscode = new OptionSetValue(910590004);
                                        service.Update(Line);
                                    }
                            }
                        }
                    }
                    #endregion
                }
                else if (Head.hil_PRType.Value == 910590000) // FOR MSL
                {
                    #region For MSL Approval
                    if (PductReqHd.hil_ApprovalStatus != null)
                    {
                        switch (PductReqHd.hil_ApprovalStatus.Value)
                        {
                            case 1:
                                Head.hil_SyncStatus = new OptionSetValue(1);
                                Head.statuscode = new OptionSetValue(910590003);
                                service.Update(Head);
                                break;
                            case 2:
                                break;
                        }
                    }
                    #endregion
                }
            }
        }
        #endregion
        public static Guid GetLineforThisHeader(IOrganizationService service, Guid HeaderId)
        {
            Guid LineId = new Guid();
            QueryByAttribute Query = new QueryByAttribute(hil_productrequest.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(false);
            Query.AddAttributeValue("hil_prheader", HeaderId);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if(Found.Entities.Count > 0)
            {
                LineId = Found.Entities[0].Id;
            }
            return LineId;
        }
        #region Conditional Assign
        public static void AssignToNextLevel(IOrganizationService service, Guid Division, int NextLevel, Guid ProdReq, Guid Header, EntityReference iFall)
        {
            if (iFall != null)
            {
                Entity SBUDiv = service.Retrieve("hil_sbubranchmapping", iFall.Id, new ColumnSet("hil_branchheaduser", "hil_nph", "hil_nsh"));
                if (NextLevel == 1)
                {
                    if (SBUDiv.Contains("hil_branchheaduser") && SBUDiv.Attributes.Contains("hil_branchheaduser"))
                    {
                        EntityReference Branch = SBUDiv.GetAttributeValue<EntityReference>("hil_branchheaduser");
                        Helper.Assign(SystemUser.EntityLogicalName, hil_productrequest.EntityLogicalName, Branch.Id, ProdReq, service);
                        Helper.Assign(SystemUser.EntityLogicalName, hil_productrequestheader.EntityLogicalName, Branch.Id, Header, service);
                    }
                }
                if (NextLevel == 2)
                {
                    if (SBUDiv.Contains("hil_nph") && SBUDiv.Attributes.Contains("hil_nph"))
                    {
                        EntityReference Branch = SBUDiv.GetAttributeValue<EntityReference>("hil_nph");
                        Helper.Assign(SystemUser.EntityLogicalName, hil_productrequest.EntityLogicalName, Branch.Id, ProdReq, service);
                        Helper.Assign(SystemUser.EntityLogicalName, hil_productrequestheader.EntityLogicalName, Branch.Id, Header, service);
                    }
                }
                if (NextLevel == 3)
                {
                    if (SBUDiv.Contains("hil_nsh") && SBUDiv.Attributes.Contains("hil_nsh"))
                    {
                        EntityReference Branch = SBUDiv.GetAttributeValue<EntityReference>("hil_nsh");
                        Helper.Assign(SystemUser.EntityLogicalName, hil_productrequest.EntityLogicalName, Branch.Id, ProdReq, service);
                        Helper.Assign(SystemUser.EntityLogicalName, hil_productrequestheader.EntityLogicalName, Branch.Id, Header, service);
                    }
                }
            }
        }
        #endregion
    }
}