using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
namespace Havells_Plugin.JobStatuses
{
    public  class StatusTransition : IPlugin
    {
        private static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                bool IsOcr = false;
                bool IsClosed = false;
                string PositionName = string.Empty;
                SystemUser iOwner = new SystemUser();
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity && context.Depth == 1)
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    #region  WORKORDER
                    if (context.PrimaryEntityName.ToLower() == msdyn_workorder.EntityLogicalName)
                    {
                        msdyn_workorder enWorkOrder = entity.ToEntity<msdyn_workorder>();
                        tracingService.Trace("1 - " + context.MessageName.ToUpper());
                        #region  UPDATE
                        //-------------->>>>>>>>IF BRANCH HEAD IS THE OWNER <<<<<<<----------------//
                        if (context.MessageName.ToUpper() == "UPDATE")
                        {
                            Entity preentity = (Entity)context.PreEntityImages["preimage"];
                            msdyn_workorder enWorkOrderPre = preentity.ToEntity<msdyn_workorder>();
                            tracingService.Trace("2 - ");
                            //ON UPDATE OF OWNER    
                            if ((enWorkOrder.OwnerId != null))
                            {
                                tracingService.Trace("3 - ");
                                if (enWorkOrder.Attributes.Contains("hil_isocr"))
                                {
                                    IsOcr = (bool)enWorkOrder["hil_isocr"];
                                }
                                if (IsOcr != true)
                                {
                                    tracingService.Trace("4 - ");
                                    iOwner = (SystemUser)service.Retrieve(SystemUser.EntityLogicalName, enWorkOrder.OwnerId.Id, new ColumnSet("positionid"));
                                    tracingService.Trace("5 - ");
                                    if (iOwner.PositionId != null)
                                    {
                                        PositionName = iOwner.PositionId.Name;
                                        if (PositionName == "BSH" || PositionName == "NSH" || PositionName == "ASH" || PositionName == "RSH" ||
                                        PositionName == "HOD" || PositionName == "CCO" || PositionName == "Call Center")
                                        {
                                            tracingService.Trace("6 - ");
                                            setWoSubStatusRecord("Pending for Allocation", entity.Id, service);
                                        }
                                        else
                                        {
                                            tracingService.Trace("7 - ");
                                            setWoSubStatusRecord("Work Allocated", entity.Id, service);
                                            //enWorkOrder["hil_assignmentdate"] = DateTime.Now.AddMinutes(330);
                                            //service.Update(enWorkOrder);
                                        }
                                    }
                                    else
                                    {
                                        tracingService.Trace("8 - ");
                                        OptionSetValue opCustomerTypeCode = GetUserAccountTypeCode(enWorkOrder.OwnerId.Id, service);
                                        tracingService.Trace("4 - " + opCustomerTypeCode);
                                        if (opCustomerTypeCode != null)
                                        {
                                            tracingService.Trace("5 - " + opCustomerTypeCode.Value);
                                            //IF BRANCH HEAD IS THE OWNER
                                            if (opCustomerTypeCode.Value == 7)// ((int)Account_CustomerTypeCode.Branch))
                                            {
                                                setWoSubStatusRecord("Pending for Allocation", entity.Id, service);
                                            }
                                            else if (opCustomerTypeCode.Value == 6)// ((int)Account_CustomerTypeCode.Franchisee))
                                            {
                                                setWoSubStatusRecord("Work Allocated", entity.Id, service);
                                                //enWorkOrder["hil_assignmentdate"] = DateTime.Now.AddMinutes(330);
                                                //service.Update(enWorkOrder);
                                            }
                                            else if (opCustomerTypeCode.Value == 9)//((int)Account_CustomerTypeCode.DirectEngineer))
                                            {
                                                setWoSubStatusRecord("Work Allocated", entity.Id, service);
                                                //enWorkOrder["hil_assignmentdate"] = DateTime.Now.AddMinutes(330);
                                                //service.Update(enWorkOrder);
                                            }
                                            tracingService.Trace("6");
                                        }
                                    }
                                }
                                else
                                {
                                    setWoSubStatusRecordJobClose("Closed", entity.Id, service);
                                    entity["hil_jobclosuredon"] = (DateTime)DateTime.Now.AddMinutes(330);
                                    entity["msdyn_timeclosed"] = (DateTime)DateTime.Now.AddMinutes(330);
                                    entity["hil_jobclosureon"] = (DateTime)DateTime.Now.AddMinutes(330);

                                    #region Added By Kuldeep Khare 25/Nov/2019 to calculate OCR Call TAT Hr & Category
                                    entity["hil_tattime"] = (DateTime)DateTime.Now.AddMinutes(330);
                                    DateTime WorkDoneOn = (DateTime)DateTime.Now.AddMinutes(330);
                                    DateTime CreatedOn = Convert.ToDateTime(enWorkOrder.CreatedOn);
                                    TimeSpan diff = WorkDoneOn - CreatedOn;
                                    double hours = diff.TotalMinutes / 60;
                                    entity["hil_tattimecalculated"] = Convert.ToDecimal(hours);
                                    EntityReference entRef = Havells_Plugin.WorkOrder.Common.TATCategory(service, hours);
                                    if (entRef != null)
                                    {
                                        entity["hil_tatcategory"] = entRef;
                                    }
                                    #endregion

                                    service.Update(entity);
                                }
                            }
                            //-------------->>>>>>>>WHEN USER MARK WORK STARTED<<<<<<<----------------//
                            else if (enWorkOrder.hil_WorkStarted != null && enWorkOrder.hil_WorkStarted.Value == true && enWorkOrderPre.hil_WorkStarted.Value == false)
                            {
                                tracingService.Trace("6");
                                if (enWorkOrder.hil_WorkStarted.Value == true)
                                {
                                    tracingService.Trace("6.1");
                                    if (enWorkOrderPre.hil_WorkStarted.Value == false)
                                    {
                                        tracingService.Trace("6.2");
                                        setWoSubStatusRecord("Work Initiated", entity.Id, service);
                                    }
                                }
                            }
                            //-------------->>>>>>>>WHEN USER MARK CALCULATE CHARGES<<<<<<<----------------//
                            else if (enWorkOrder.hil_CloseTicket != null && enWorkOrder.hil_CloseTicket.Value == true && enWorkOrderPre.hil_CloseTicket.Value == false)
                            {
                                tracingService.Trace("7");
                                if (enWorkOrder.hil_CloseTicket.Value == true)
                                {
                                    tracingService.Trace("7.1");
                                    if (enWorkOrderPre.hil_CloseTicket.Value == false)
                                    {
                                        tracingService.Trace("7.2");
                                        msdyn_workorder msWo = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, entity.Id, new Microsoft.Xrm.Sdk.Query.ColumnSet("hil_closureremarks"));
                                        if (msWo.hil_ClosureRemarks == null)
                                            throw new InvalidPluginExecutionException("----->>>Can't close service ticket without filling Closure Reamrks<<<-----");
                                        #region  Update Job Closure Time on Job
                                        msdyn_workorder upWO = new msdyn_workorder();
                                        upWO.msdyn_workorderId = entity.Id;
                                        upWO["hil_jobclosureon"] = DateTime.Now.AddMinutes(330);
                                        entity["msdyn_timeclosed"] = (DateTime)DateTime.Now.AddMinutes(330);
                                        //upWO.hil_JobClosuredon = DateTime.Now;
                                        service.Update(upWO);
                                        tracingService.Trace("7.2.1");
                                        #endregion
                                        setWoSubStatusRecordJobClose("Closed", entity.Id, service);
                                        //setWoSubStatusRecordJobClose("Test Close", entity.Id, service);

                                        //
                                        tracingService.Trace("7.3");
                                    }
                                }
                            }

                            //-------------->>>>>>>>WHEN USER MARK ASSIGN TO ME<<<<<<<----------------//
                            else if (enWorkOrder.hil_Assigntome != null && enWorkOrder.hil_Assigntome.Value == true && enWorkOrderPre.hil_Assigntome.Value == false)
                            {
                                tracingService.Trace("8");
                                if (enWorkOrder.hil_Assigntome.Value == true)
                                {

                                    tracingService.Trace("8.1");
                                    if (enWorkOrderPre.hil_Assigntome.Value == false)
                                    {
                                        tracingService.Trace("8.2");
                                        setWoSubStatusRecord("Work Allocated", entity.Id, service);
                                    }
                                }
                            }
                            //-------------->>>>>>>>WHEN USER CLICK CREATE PO BUTTON<<<<<<<----------------//
                            else if (enWorkOrder.hil_FlagPO != null && enWorkOrder.hil_FlagPO.Value == true && enWorkOrderPre.hil_FlagPO.Value == false)
                            {

                                tracingService.Trace("9.1");
                                if (enWorkOrder.hil_FlagPO.Value == true)
                                {
                                    tracingService.Trace("9.2");
                                    if (enWorkOrderPre.hil_FlagPO.Value == false)
                                    {
                                        tracingService.Trace("9.3");
                                        setWoSubStatusRecord("Part PO Created", entity.Id, service);
                                    }
                                }
                            }
                            //-------------->>>>>>>>WHEN USER MARK CLOSE TICKET<<<<<<<----------------//
                            else if (enWorkOrder.hil_CalculateCharges != null && enWorkOrder.hil_CalculateCharges.Value == true && enWorkOrderPre.hil_CalculateCharges.Value == false)
                            {
                                tracingService.Trace("10");
                                if (enWorkOrder.hil_CalculateCharges.Value == true)
                                {
                                    tracingService.Trace("10.1");
                                    if (enWorkOrderPre.hil_CalculateCharges.Value == false)
                                    {
                                        tracingService.Trace("10.2");
                                        setWoSubStatusRecord("Work Done", entity.Id, service);
                                    }
                                }
                            }
                            tracingService.Trace("11");
                        }
                        #endregion
                        #region  CREATE
                        if (context.MessageName.ToUpper() == "CREATE")
                        {
                            if (enWorkOrder.OwnerId != null)
                            {
                                OptionSetValue opCustomerTypeCode = GetUserAccountTypeCode(enWorkOrder.OwnerId.Id, service);
                                if (opCustomerTypeCode != null)
                                {
                                    if(enWorkOrder.Attributes.Contains("hil_isocr"))
                                    {
                                        IsOcr = (bool)enWorkOrder["hil_isocr"];
                                    }
                                    if(enWorkOrder.Attributes.Contains("hil_ifclosedjob"))
                                    {
                                        IsClosed = (bool)enWorkOrder["hil_ifclosedjob"];
                                    }
                                    if(IsOcr != true && IsClosed != true)
                                    {
                                        //IF BRANCH HEAD IS THE OWNER
                                        if (opCustomerTypeCode.Value == ((int)Account_CustomerTypeCode.Branch))
                                        {
                                            setWoSubStatusRecord("Pending for Allocation", entity.Id, service);
                                        }
                                        else if (opCustomerTypeCode.Value == ((int)Account_CustomerTypeCode.Franchisee) || opCustomerTypeCode.Value == ((int)Account_CustomerTypeCode.DirectEngineer))
                                        {
                                            setWoSubStatusRecord("Work Allocated", entity.Id, service);
                                        }
                                    }
                                    else
                                    {
                                        setWoSubStatusRecordJobClose("Closed", entity.Id, service);
                                        entity["hil_jobclosuredon"] = (DateTime)DateTime.Now.AddMinutes(330);
                                        entity["msdyn_timeclosed"] = (DateTime)DateTime.Now.AddMinutes(330);
                                        entity["hil_jobclosureon"] = (DateTime)DateTime.Now.AddMinutes(330);

                                        #region Added By Kuldeep Khare 25/Nov/2019 to calculate OCR Call TAT Hr & Category
                                        entity["hil_tattime"] = (DateTime)DateTime.Now.AddMinutes(330);
                                        entity["hil_tattimecalculated"] = Convert.ToDecimal(0);
                                        EntityReference entRef = Havells_Plugin.WorkOrder.Common.TATCategory(service, 0);
                                        if (entRef != null)
                                        {
                                            entity["hil_tatcategory"] = entRef;
                                        }
                                        #endregion

                                        service.Update(entity);
                                    }
                                }
                            }
                        }
                        #endregion
                    }
                    #endregion
                    #region  GRNLine
                    if (context.PrimaryEntityName.ToLower() == hil_grnline.EntityLogicalName)
                    {
                        if (context.MessageName.ToUpper() == "UPDATE")
                        {
                              UpdateWOStatusonGrnLine(entity, service);
                        }
                    }
                    #endregion
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        public static void UpdateWOStatusonGrnLine(Entity entity,IOrganizationService service)
        {
            hil_grnline iGRN = (hil_grnline)service.Retrieve(hil_grnline.EntityLogicalName, entity.Id, new ColumnSet("hil_productrequest"));
            hil_productrequest iReq = (hil_productrequest)service.Retrieve(hil_productrequest.EntityLogicalName, iGRN.hil_ProductRequest.Id, new ColumnSet("hil_job"));
            if(iReq.hil_Job != null)
            {
                try
                {
                    using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                    {
                        hil_grnline enGrnLine = entity.ToEntity<hil_grnline>();
                        //execution order 2
                        if ((enGrnLine.hil_ApproverStatus != null && enGrnLine.hil_ApproverStatus.Value == 1)
                            || (enGrnLine.statuscode != null && enGrnLine.statuscode.Value == 910590000))//Approval Done or Status Confirmed
                        {
                            int iReceivedQtyGood = 0;
                            int iRequestedQty = 0;
                            Guid fsWorkOrderId = Guid.Empty;

                            #region GetAllGrnLinesQty
                            var obj = from _GrnLine in orgContext.CreateQuery<hil_grnline>()
                                      join _PO in orgContext.CreateQuery<hil_productrequest>()
                                      on _GrnLine.hil_ProductRequest.Id equals _PO.hil_productrequestId.Value
                                      join _woprod in orgContext.CreateQuery<msdyn_workorderproduct>()
                                      on _PO.hil_productrequestId.Value equals _woprod.hil_purchaseorder.Id
                                      join _WorkOrder in orgContext.CreateQuery<msdyn_workorder>()
                                      on _woprod.msdyn_WorkOrder.Id equals _WorkOrder.msdyn_workorderId.Value
                                      where _GrnLine.hil_grnlineId.Value == enGrnLine.Id
                                      select new
                                      {
                                          _GrnLine.hil_GoodQuantity
                                          ,
                                          _WorkOrder.msdyn_workorderId
                                          ,
                                          _GrnLine.hil_ApprovedDefectiveQuantity
                                          ,
                                          _GrnLine.hil_DefectiveQuantity
                                      };
                            foreach (var iobj in obj)
                            {
                                if (iobj.hil_GoodQuantity != null)
                                {
                                    iReceivedQtyGood += iobj.hil_GoodQuantity.Value;
                                }
                                if (iobj.hil_DefectiveQuantity != null && iobj.hil_ApprovedDefectiveQuantity != null)
                                {
                                    int temp = iobj.hil_DefectiveQuantity.Value - iobj.hil_ApprovedDefectiveQuantity.Value;
                                    iRequestedQty += temp;
                                }
                                fsWorkOrderId = iobj.msdyn_workorderId.Value;
                            }
                            #endregion
                            #region GetPORequested
                            var obj1 = from _woProduct in orgContext.CreateQuery<msdyn_workorderproduct>()
                                       join _PO in orgContext.CreateQuery<hil_productrequest>()
                                       on _woProduct.hil_purchaseorder.Id equals _PO.hil_productrequestId.Value
                                       where _woProduct.msdyn_WorkOrder.Id == fsWorkOrderId
                                       select new
                                       {
                                           _PO.hil_Quantity
                                       };
                            foreach (var iobj1 in obj1)
                            {
                                if (iobj1.hil_Quantity != null)
                                    iRequestedQty += iobj1.hil_Quantity.Value;
                            }
                            #endregion

                            //if whole quantity received
                            if (iReceivedQtyGood == iRequestedQty && iRequestedQty != 0)
                            {
                                setWoSubStatusRecord("Parts Fulfilled", fsWorkOrderId, service);
                            }
                            else
                            {
                                setWoSubStatusRecord("Work Initiated", fsWorkOrderId, service);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException(ex.Message);
                }
            }
        }

        public static OptionSetValue GetUserAccountTypeCode(Guid fsUserId, IOrganizationService service)
        {
            try
            {
                OptionSetValue opCustomerTypeCode = null;
                using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                {

                    var obj = (from _User in orgContext.CreateQuery<SystemUser>()
                               join _Account in orgContext.CreateQuery<Account>()

                               on _User.hil_Account.Id equals _Account.AccountId.Value
                               join _Position in orgContext.CreateQuery<Position>()
                               on _User.PositionId.Id equals _Position.PositionId.Value
                               where _User.SystemUserId.Value== fsUserId
                               select new { _Account.CustomerTypeCode, _Position.Name }).Take(1);

                    foreach (var iobj in obj)
                    {
                        if (iobj.CustomerTypeCode != null)
                        {
                            opCustomerTypeCode = iobj.CustomerTypeCode;
                        }
                    }
                }
                return opCustomerTypeCode;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        public static void setWoSubStatusRecord(String sSubStatus,Guid fsWorkOrderId, IOrganizationService service)
        {
            try
            {
                msdyn_workorder enJob = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, fsWorkOrderId, new ColumnSet("msdyn_substatus", "hil_isocr"));
                if (enJob.msdyn_SubStatus != null)
                {
                    if(enJob.msdyn_SubStatus.Name != "Closed" ||enJob.msdyn_SubStatus.Name != "Canceled")
                    {
                        if(sSubStatus == "Closed" && enJob.msdyn_SubStatus.Name == "Work Alocated")
                        {
                            if(enJob.Contains("hil_isocr") && enJob.Attributes.Contains("hil_isocr"))
                            {
                                bool bIsOCR = enJob.GetAttributeValue<bool>("hil_isocr");
                                if(bIsOCR != true)
                                {
                                    throw new InvalidPluginExecutionException("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX - STATUS CAN'T BE CHANGED - XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX");
                                }
                                else
                                {
                                    ChangeStatus(sSubStatus, fsWorkOrderId, service);
                                }
                            }
                            else
                            {
                                throw new InvalidPluginExecutionException("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX - STATUS CAN'T BE CHANGED - XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX");
                            }
                        }
                        else
                        {
                            ChangeStatus(sSubStatus, fsWorkOrderId, service);
                        }
                    }
                }
                else
                {
                    ChangeStatus(sSubStatus, fsWorkOrderId, service);
                }
                
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        public static void setWoSubStatusRecordJobClose(String sSubStatus, Guid fsWorkOrderId, IOrganizationService service)
        {
            try
            {
                tracingService.Trace("7.999");
                using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                {
                    var obj = from _workOrderSubStatus in orgContext.CreateQuery<msdyn_workordersubstatus>()
                              where _workOrderSubStatus.msdyn_name == sSubStatus
                              select new
                              {
                                  _workOrderSubStatus.msdyn_workordersubstatusId
                                  ,
                                  _workOrderSubStatus.msdyn_SystemStatus
                              };
                    foreach (var iobj in obj)
                    {
                        tracingService.Trace("7.998");
                        msdyn_workorder upWorkOrder = new msdyn_workorder();
                        upWorkOrder.msdyn_workorderId = fsWorkOrderId;
                        upWorkOrder.msdyn_SubStatus = new EntityReference(msdyn_workordersubstatus.EntityLogicalName, iobj.msdyn_workordersubstatusId.Value);
                        if (iobj.msdyn_SystemStatus != null)
                        {
                            upWorkOrder.msdyn_SystemStatus = iobj.msdyn_SystemStatus;
                        }
                        service.Update(upWorkOrder);
                        tracingService.Trace("7.997");
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        public static void ChangeStatus(String sSubStatus, Guid fsWorkOrderId, IOrganizationService service)
        {
            using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
            {
                var obj = from _workOrderSubStatus in orgContext.CreateQuery<msdyn_workordersubstatus>()
                          where _workOrderSubStatus.msdyn_name == sSubStatus
                          select new
                          {
                              _workOrderSubStatus.msdyn_workordersubstatusId
                              ,
                              _workOrderSubStatus.msdyn_SystemStatus
                          };
                foreach (var iobj in obj)
                {
                    msdyn_workorder upWorkOrder = new msdyn_workorder();
                    upWorkOrder.msdyn_workorderId = fsWorkOrderId;
                    upWorkOrder.msdyn_SubStatus = new EntityReference(msdyn_workordersubstatus.EntityLogicalName, iobj.msdyn_workordersubstatusId.Value);
                    if (iobj.msdyn_SystemStatus != null)
                    {
                        upWorkOrder.msdyn_SystemStatus = iobj.msdyn_SystemStatus;
                    }
                    service.Update(upWorkOrder);
                }
            }
        }
    }
}


