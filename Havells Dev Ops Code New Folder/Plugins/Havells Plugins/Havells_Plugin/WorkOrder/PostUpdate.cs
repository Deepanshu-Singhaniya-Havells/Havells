using Havells_Plugin.SAWActivity;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;

namespace Havells_Plugin.WorkOrder
{

    public class PostUpdate : IPlugin
    {
        static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #region MainRegion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == "msdyn_workorder" && context.MessageName.ToUpper() == "UPDATE" && context.Depth == 1)
                {
                    tracingService.Trace("2");
                    Entity entity = (Entity)context.InputParameters["Target"];
                    msdyn_workorder enWorkorder = entity.ToEntity<msdyn_workorder>();
                    msdyn_workorder preImageWO = ((Entity)context.PreEntityImages["image"]).ToEntity<msdyn_workorder>();
                    tracingService.Trace("233");
                    #region CreatePOForWOProducts
                    WorkOrderPO(entity, service);
                    //Added By Kuldeep Khare 13/Nov/2019 KKG Audit Failure -> Create new Child Job
                    if (entity.Attributes.Contains("hil_productsubcategory") || entity.Attributes.Contains("hil_natureofcomplaint"))
                    {
                        try
                        {
                            Entity _jobData = service.Retrieve(msdyn_workorder.EntityLogicalName, entity.Id, new ColumnSet("hil_productsubcategory", "hil_natureofcomplaint"));

                            if (_jobData.Contains("hil_productsubcategory") && _jobData.Contains("hil_natureofcomplaint"))
                            {
                                Guid _productsubCategoryId = _jobData.GetAttributeValue<EntityReference>("hil_productsubcategory").Id;
                                Guid _natureofcomplaintId = _jobData.GetAttributeValue<EntityReference>("hil_natureofcomplaint").Id;

                                QueryExpression qryExp = new QueryExpression("hil_natureofcomplaint");
                                qryExp.ColumnSet = new ColumnSet(false);
                                qryExp.Criteria.AddCondition(new ConditionExpression("hil_relatedproduct", ConditionOperator.Equal, _productsubCategoryId));
                                qryExp.Criteria.AddCondition(new ConditionExpression("hil_natureofcomplaintid", ConditionOperator.Equal, _natureofcomplaintId));
                                qryExp.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                                EntityCollection entCol = service.RetrieveMultiple(qryExp);
                                if (entCol.Entities.Count == 0)
                                {
                                    throw new InvalidPluginExecutionException("Invalid Nature of Complaint.");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            throw new InvalidPluginExecutionException(ex.Message);
                        }
                    }
                    if (entity.Attributes.Contains("hil_createchildjob"))
                    {
                        Boolean temp = entity.GetAttributeValue<Boolean>("hil_createchildjob");
                        if (temp == true)
                        {
                            CreateChildJob(service, preImageWO);
                        }
                    }
                    if (entity.Attributes.Contains("hil_calculatecharges"))
                    {
                        Boolean temp = entity.GetAttributeValue<Boolean>("hil_calculatecharges");
                        if (temp == true)
                        {
                            UpdateJobWarrantyDetails(entity.Id, service);
                            bool IfPartsEstimated = false;
                            if (IfPartsEstimated == false)
                            {
                                IfPartsEstimated = CheckIfReplacedPartNullOrNotAvailable(service, entity.Id);
                                if (IfPartsEstimated == false)
                                {
                                    IfPartsEstimated = CheckIfActionsUsed(service, entity.Id);
                                    if (IfPartsEstimated == false)
                                    {
                                        tracingService.Trace("23");
                                        CalculateCharges(service, entity.Id, preImageWO);
                                        tracingService.Trace("3");
                                        CalculateTAT(service, entity.Id);
                                    }
                                    else
                                    {
                                        throw new InvalidPluginExecutionException("XXXXXXXXXXXXXXXXXXXXXXXXXXXX - Atlease One Action is Required Before Marking Work Done - XXXXXXXXXXXXXXXXXXXXXXXXXXXX");
                                    }
                                }
                                else
                                {
                                    throw new InvalidPluginExecutionException("XXXXXXXXXXXXXXXXXXXXXXXXXXXX - Please Close Open Job Products or Services or PO Still in Process - XXXXXXXXXXXXXXXXXXXXXXXXXXXX");
                                }
                            }
                            else
                            {
                                throw new InvalidPluginExecutionException("XXXXXXXXXXXXXXXXXXXXXXXXXXXX - Please Close Open Job Products or Services or PO Still in Process - XXXXXXXXXXXXXXXXXXXXXXXXXXXX");
                            }
                            //Added by Kuldeep Khare 15/Nov/2019 as discussed with Paliwal sir, to update Asset First Installation details on Job Done
                            //UpdateFirstInstallationInformationOnAsset(service, entity.Id);

                        }
                    }
                    if (entity.Attributes.Contains("hil_estimatecharges"))
                    {
                        Boolean temp = entity.GetAttributeValue<Boolean>("hil_estimatecharges");
                        if (temp)
                        {
                            EstimateCharges(service, entity.Id);
                            CreateEstimateHeader(service, entity.Id);
                        }
                    }

                    if (entity.Attributes.Contains("hil_closeticket"))
                    {
                        bool _cVal = (bool)entity["hil_closeticket"];
                        if (_cVal == true)
                        {
                            msdyn_workorder _enJobTemp = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, entity.Id, new ColumnSet("hil_callsubtype", "hil_receiptamount", "hil_modeofpayment", "msdyn_substatus", "hil_callsubtype"));
                            //AMC Call
                            if (_enJobTemp != null && _enJobTemp.GetAttributeValue<EntityReference>("hil_callsubtype").Id == new Guid(JobCallSubType.AMCCall))
                            {
                                //QueryExpression expJobProduct = new QueryExpression("msdyn_workorderproduct");
                                //expJobProduct.ColumnSet = new ColumnSet("msdyn_product");
                                //expJobProduct.Criteria = new FilterExpression(LogicalOperator.And);
                                //expJobProduct.Criteria.AddCondition("msdyn_workorder", ConditionOperator.Equal, entity.Id);
                                //LinkEntity EntityA = new LinkEntity("msdyn_workorderproduct", "hil_productcatalog", "msdyn_product", "hil_productcode", JoinOperator.LeftOuter);
                                //EntityA.Columns = new ColumnSet("hil_amctandc");
                                //EntityA.EntityAlias = "Prod";
                                //expJobProduct.LinkEntities.Add(EntityA);
                                //LinkEntity EntityB = new LinkEntity("msdyn_workorderproduct", "product", "msdyn_product", "productid", JoinOperator.Inner);
                                //EntityB.Columns = new ColumnSet(false);
                                //EntityB.LinkCriteria = new FilterExpression(LogicalOperator.And);
                                //EntityB.LinkCriteria.AddCondition("hil_hierarchylevel", ConditionOperator.Equal, 910590001); //AMC Product
                                //EntityB.LinkCriteria.AddCondition("msdyn_fieldserviceproducttype", ConditionOperator.Equal, 690970001); //Non-Inventory
                                //EntityB.EntityAlias = "ProdAMC";
                                //expJobProduct.LinkEntities.Add(EntityB);
                                //EntityCollection entCollJobProd = service.RetrieveMultiple(expJobProduct);

                                QueryExpression expJobProduct = new QueryExpression("msdyn_workorderproduct");
                                expJobProduct.ColumnSet = new ColumnSet("msdyn_product");
                                expJobProduct.Criteria = new FilterExpression(LogicalOperator.And);
                                expJobProduct.Criteria.AddCondition("msdyn_workorder", ConditionOperator.Equal, entity.Id);
                                expJobProduct.Criteria.AddCondition("hil_markused", ConditionOperator.Equal, true);
                                expJobProduct.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                                LinkEntity EntityA = new LinkEntity("msdyn_workorderproduct", "hil_productcatalog", "hil_replacedpart", "hil_productcode", JoinOperator.LeftOuter);
                                EntityA.Columns = new ColumnSet("hil_amctandc");
                                EntityA.EntityAlias = "Prod";
                                EntityA.LinkCriteria.AddCondition("statecode", ConditionOperator.Equal, 0); // Active Product Catalog Line
                                expJobProduct.LinkEntities.Add(EntityA);
                                LinkEntity EntityB = new LinkEntity("msdyn_workorderproduct", "product", "hil_replacedpart", "productid", JoinOperator.Inner);
                                EntityB.Columns = new ColumnSet(false);
                                EntityB.LinkCriteria = new FilterExpression(LogicalOperator.And);
                                EntityB.LinkCriteria.AddCondition("hil_hierarchylevel", ConditionOperator.Equal, 910590001); //AMC Product
                                EntityB.LinkCriteria.AddCondition("msdyn_fieldserviceproducttype", ConditionOperator.Equal, 690970001); //Non-Inventory
                                EntityB.EntityAlias = "ProdAMC";
                                expJobProduct.LinkEntities.Add(EntityB);
                                EntityCollection entCollJobProd = service.RetrieveMultiple(expJobProduct);
                                if (entCollJobProd.Entities.Count == 0)
                                {
                                    throw new InvalidPluginExecutionException("Atlease One AMC Product is Required Before Marking Job Close.");
                                }
                                else
                                {
                                    if (!entCollJobProd.Entities[0].Contains("Prod.hil_amctandc"))
                                    {
                                        throw new InvalidPluginExecutionException("Warranty Description is not defined in selected AMC Product Calatog setup.");
                                    }
                                }
                                if (_enJobTemp.Contains("hil_receiptamount"))
                                {
                                    if (_enJobTemp.GetAttributeValue<Money>("hil_receiptamount").Value == 0)
                                        throw new InvalidPluginExecutionException("Receipt Amount can't be Zero in AMC Call.");
                                }
                                else
                                {
                                    throw new InvalidPluginExecutionException("Receipt Amount can't be balnk/zero in AMC Call.");
                                }
                            }
                            //InstallationDateOnAsset(service, entity.Id);
                            UpdateFirstInstallationInformationOnAsset(service, entity.Id);
                            UpdateAssetAndWarrantyDetailsOnJob(service, entity);
                            CreateWarrantyForReplacedPart(service, entity.Id);
                            UpdateJobclosureSourceforMobileApp(service, entity.Id);
                            UpdateJobUpcountry(service, entity.Id);
                            //UpdateClaimParams(service, entity.Id);
                            //======================START Update Job Closure Details Added By Kuldeep Khare 31/Dec/2019 =============
                            //msdyn_workorder woEntity = new msdyn_workorder();//new msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, TicketId, new ColumnSet(false));
                            //woEntity.Id = entity.Id;
                            //woEntity.msdyn_TimeClosed = (DateTime)DateTime.Now.AddMinutes(330);
                            //woEntity.msdyn_ClosedBy = new EntityReference("systemuser", context.UserId);
                            //service.Update(woEntity);
                            //====================== END ============================================================================
                            #region Post Closure SAW Activity Approvals
                            msdyn_workorder _enJob1 = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, entity.Id, new ColumnSet("msdyn_customerasset", "msdyn_timeclosed", "hil_typeofassignee", "hil_isgascharged", "hil_laborinwarranty", "hil_callsubtype", "hil_productsubcategory", "hil_customerref", "createdon", "hil_isocr", "hil_claimstatus"));
                            if (_enJob1 != null)
                            {
                                bool _isOCR = false;
                                bool _gasCharged = false;
                                bool _laborInWarranty = false;
                                bool _underReview = false;
                                OptionSetValue optVal = null;
                                EntityReference erTypeOfAssignee = null;

                                if (_enJob1.Attributes.Contains("hil_typeofassignee"))
                                {
                                    erTypeOfAssignee = _enJob1.GetAttributeValue<EntityReference>("hil_typeofassignee");
                                }
                                if (_enJob1.Attributes.Contains("hil_claimstatus"))
                                {
                                    optVal = _enJob1.GetAttributeValue<OptionSetValue>("hil_claimstatus");
                                }
                                if (_enJob1.Attributes.Contains("hil_isocr"))
                                {
                                    _isOCR = (bool)_enJob1["hil_isocr"];
                                }
                                if (_enJob1.Attributes.Contains("hil_isgascharged"))
                                {
                                    _gasCharged = (bool)_enJob1["hil_isgascharged"];
                                }
                                if (_enJob1.Attributes.Contains("hil_laborinwarranty"))
                                {
                                    _laborInWarranty = (bool)_enJob1["hil_laborinwarranty"];
                                }
                                if (!_isOCR)
                                {//_laborInWarranty && 
                                    if (erTypeOfAssignee.Id != new Guid("7D1ECBAB-1208-E911-A94D-000D3AF0694E")) // Labor InWarranty and Type of Assignee !=DSE
                                    {
                                        #region Gas Charged Approval
                                        if (_gasCharged)
                                        {
                                            CommonLib obj = new CommonLib();
                                            CommonLib objReturn = obj.CreateSAWActivity(entity.Id, 0, SAWCategoryConst._GasChargePostAudit, service, "", null);
                                            if (objReturn.statusRemarks == "OK")
                                            {
                                                _underReview = true;
                                            }
                                        }
                                        #endregion

                                        #region RepeatRepair Approval
                                        DateTime _createdOn = _enJob1.GetAttributeValue<DateTime>("createdon").AddDays(-15);
                                        DateTime _ClosedOn = DateTime.Now.AddDays(-15);
                                        string _strCreatedOn = _createdOn.Year.ToString() + "-" + _createdOn.Month.ToString() + "-" + _createdOn.Day.ToString();
                                        string _strClosedOn = _ClosedOn.Year.ToString() + "-" + _ClosedOn.Month.ToString() + "-" + _ClosedOn.Day.ToString();
                                        EntityCollection entCol;

                                        //Callsubtype{8D80346B-3C0B-E911-A94E-000D3AF06CD} Dealer Stock Repair
                                        //JobStatus: {1727FA6C-FA0F-E911-A94E-000D3AF060A1}-Closed,2927FA6C-FA0F-E911-A94E-000D3AF060A1-Workdone,7E85074C-9C54-E911-A951-000D3AF0677F-workdone SMS
                                        string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                        <entity name='msdyn_workorder'>
                                        <attribute name='msdyn_name' />
                                        <attribute name='createdon' />
                                        <attribute name='hil_productsubcategory' />
                                        <attribute name='hil_customerref' />
                                        <attribute name='msdyn_customerasset' />
                                        <attribute name='hil_callsubtype' />
                                        <attribute name='msdyn_workorderid' />
                                        <attribute name='msdyn_timeclosed' />
                                        <attribute name='msdyn_closedby' />
                                        <order attribute='createdon' descending='true' />
                                        <filter type='and'>
                                            <condition attribute='hil_isocr' operator='ne' value='1' />
                                            <condition attribute='hil_typeofassignee' operator='ne' value='{7D1ECBAB-1208-E911-A94D-000D3AF0694E}' />
                                            <condition attribute='msdyn_workorderid' operator='ne' value='" + entity.Id + @"' />
                                            <condition attribute='hil_customerref' operator='eq' value='" + _enJob1.GetAttributeValue<EntityReference>("hil_customerref").Id + @"' />
                                            <condition attribute='hil_callsubtype' operator='eq' value='" + _enJob1.GetAttributeValue<EntityReference>("hil_callsubtype").Id + @"' />
                                            <condition attribute='hil_callsubtype' operator='ne' value='{8D80346B-3C0B-E911-A94E-000D3AF06CD4}' />
                                            <condition attribute='hil_productsubcategory' operator='eq' value='" + _enJob1.GetAttributeValue<EntityReference>("hil_productsubcategory").Id + @"' />
                                            <filter type='or'>
                                                <condition attribute='msdyn_timeclosed' operator='on-or-after' value='" + _strCreatedOn + @"' />
                                                <condition attribute='msdyn_timeclosed' operator='on-or-after' value='" + _strClosedOn + @"' />
                                            </filter>
                                            <condition attribute='msdyn_substatus' operator='in'>
                                            <value>{1727FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                                            <value>{2927FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                                            <value>{7E85074C-9C54-E911-A951-000D3AF0677F}</value>
                                            </condition>
                                        </filter>
                                        </entity>
                                        </fetch>";
                                        entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                                        if (entCol.Entities.Count > 0)
                                        {
                                            //string _remarks = "Old Job# " + entCol.Entities[0].GetAttributeValue<string>("msdyn_name");
                                            //CommonLib obj = new CommonLib();
                                            //CommonLib objReturn = obj.CreateSAWActivity(entity.Id, 0, SAWCategoryConst._RepeatRepair, service, _remarks, entCol.Entities[0].ToEntityReference());
                                            string _remarks = string.Empty;
                                            string _remarks1 = string.Empty;
                                            EntityReference _entref = null;
                                            foreach (Entity ent in entCol.Entities)
                                            {
                                                _entref = ent.ToEntityReference();
                                                if (ent.GetAttributeValue<EntityReference>("msdyn_customerasset").Id == _enJob1.GetAttributeValue<EntityReference>("msdyn_customerasset").Id)
                                                {
                                                    _remarks += ent.GetAttributeValue<string>("msdyn_name") + ",";
                                                }
                                                else
                                                {
                                                    _remarks1 += ent.GetAttributeValue<string>("msdyn_name") + ",";
                                                }
                                            }
                                            _remarks = ((_remarks == string.Empty ? "" : "Repeated Jobs with Same Serial Number: " + _remarks + ":\n") + (_remarks1 == string.Empty ? "" : "Repeated Jobs with Same Product Subcategory: " + _remarks1 + ":")).Replace(",:", "");
                                            CommonLib obj = new CommonLib();
                                            CommonLib objReturn = obj.CreateSAWActivity(_enJob1.Id, 0, SAWCategoryConst._RepeatRepair, service, _remarks, _entref);
                                            if (objReturn.statusRemarks == "OK")
                                            {
                                                _underReview = true;
                                            }
                                        }
                                        #endregion

                                        if (_underReview)
                                        {
                                            if (optVal != null && optVal.Value != 3)
                                            {
                                                Entity Ticket = new Entity("msdyn_workorder");
                                                Ticket.Id = entity.Id;
                                                Ticket["hil_claimstatus"] = new OptionSetValue(1); //Claim Under Review
                                                service.Update(Ticket);
                                            }
                                        }
                                        else
                                        {
                                            QueryExpression Query;
                                            EntityCollection entcoll;
                                            Query = new QueryExpression("hil_sawactivity");
                                            Query.ColumnSet = new ColumnSet("hil_sawactivityid");
                                            Query.Criteria = new FilterExpression(LogicalOperator.And);
                                            Query.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, entity.Id);
                                            entcoll = service.RetrieveMultiple(Query);
                                            if (entcoll.Entities.Count == 0)
                                            {
                                                Entity Ticket = new Entity("msdyn_workorder");
                                                Ticket.Id = entity.Id;
                                                Ticket["hil_claimstatus"] = new OptionSetValue(4); //Claim Approved
                                                service.Update(Ticket);
                                            }
                                        }
                                    }
                                }
                            }
                            #endregion
                            //if (_enJobTemp.GetAttributeValue<EntityReference>("msdyn_substatus").Name.ToUpper() == "CLOSED")
                            //{
                            if (_enJobTemp.Attributes.Contains("hil_modeofpayment"))
                            {
                                if (_enJobTemp.GetAttributeValue<OptionSetValue>("hil_modeofpayment").Value == 6)
                                {
                                    if (_enJobTemp.GetAttributeValue<EntityReference>("hil_callsubtype").Id == new Guid("55a71a52-3c0b-e911-a94e-000d3af06cd4")) // AMC Jobs
                                    {
                                        if (!Common.ValidateOnlinePayment(service, enWorkorder.Id))
                                        {
                                            throw new InvalidPluginExecutionException("Payment for this AMC is pending, please cross check with customer & try again.");
                                        }
                                    }
                                }
                            }
                            //}
                        }
                    }

                    if (entity.Attributes.Contains("hil_isgascharged"))
                    {
                        bool _gasCharged = (bool)entity["hil_isgascharged"];
                        if (_gasCharged)
                        {
                            CreateSAWActivity(entity.Id, service);
                        }
                    }
                    if (entity.Attributes.Contains("hil_kkgprovided"))
                    {
                        bool _kVal = (bool)entity["hil_kkgprovided"];
                        if (_kVal == false)
                        {
                            CreatePhoneCall(service, entity.Id);
                        }
                    }

                    if (entity.Attributes.Contains("hil_claimstatus"))
                    {
                        OptionSetValue temp = entity.GetAttributeValue<OptionSetValue>("hil_claimstatus");
                        if (temp.Value == 3) //Rejected
                        {
                            Havells_Plugin.PerformaInvoice.ClaimOperations clmOps = new Havells_Plugin.PerformaInvoice.ClaimOperations();

                            QueryExpression qrExp = new QueryExpression("hil_claimline");
                            qrExp.ColumnSet = new ColumnSet("hil_claimheader");
                            qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                            qrExp.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, entity.Id);
                            EntityCollection entCol = service.RetrieveMultiple(qrExp);
                            if (entCol.Entities.Count > 0)
                            {
                                clmOps.UpdatePerformaInvoice(service, entCol.Entities[0].GetAttributeValue<EntityReference>("hil_claimheader").Id);
                            }
                        }
                    }
                    DeLinkAssociatedPO(service, entity);
                    #endregion
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
            #endregion
        }

        public static void CreateKKGAuditFailedSAWActivity(Guid JobId, IOrganizationService _service)
        {
            string _strCreatedOn = string.Empty;
            msdyn_workorder _enJob = (msdyn_workorder)_service.Retrieve(msdyn_workorder.EntityLogicalName, JobId, new ColumnSet("msdyn_substatus", "hil_newjobid"));
            bool _underReview = false;

            if (_enJob != null && _enJob.Attributes.Contains("hil_newjobid"))
            {
                string _remarks = "KKG Audit failed, child job created # " + _enJob.GetAttributeValue<EntityReference>("hil_newjobid").Name + ". Please verify.";
                CommonLib obj = new CommonLib();
                CommonLib objReturn = obj.CreateSAWActivity(JobId, 0, SAWCategoryConst._KKGFailureReview, _service, _remarks, _enJob.GetAttributeValue<EntityReference>("hil_newjobid"));
                if (objReturn.statusRemarks == "OK")
                {
                    _underReview = true;
                }

                if (_underReview)
                {
                    Entity Ticket = new Entity("msdyn_workorder");
                    Ticket.Id = JobId;
                    Ticket["hil_claimstatus"] = new OptionSetValue(6); //Abnormal Penalty Under Review
                    _service.Update(Ticket);
                }
            }
        }

        public void UpdateJobWarrantyDetails(Guid _jobGuId, IOrganizationService _service)
        {
            DateTime _jobCreatedOn = new DateTime(1900, 1, 1);
            DateTime _assetPurchaseDate = new DateTime(1900, 1, 1);
            DateTime _unitWarrStartDate = new DateTime(1900, 1, 1);
            QueryExpression qrExp;
            EntityCollection entCol;
            EntityReference _unitWarranty = null;
            Guid _warrantyTemplateId = Guid.Empty;
            bool laborInWarranty = false;
            int _jobWarrantyStatus = 2; //OutWarranty
            int _jobWarrantySubStatus = 0;
            int _warrantyTempType = 0;
            double _jobMonth = 0;
            bool SparePartUsed = false;
            OptionSetValue _opChannelPartnerCategory = null;

            Entity entJob = _service.Retrieve(msdyn_workorder.EntityLogicalName, _jobGuId, new ColumnSet("hil_owneraccount", "createdon", "msdyn_customerasset"));
            if (entJob != null)
            {
                if (entJob.Attributes.Contains("hil_owneraccount"))
                {
                    Entity entTemp = _service.Retrieve(Account.EntityLogicalName, entJob.GetAttributeValue<EntityReference>("hil_owneraccount").Id, new ColumnSet("hil_category"));
                    if (entTemp != null)
                    {
                        if (entTemp.Attributes.Contains("hil_category"))
                        {
                            _opChannelPartnerCategory = entTemp.GetAttributeValue<OptionSetValue>("hil_category");
                        }
                    }
                }
                _jobCreatedOn = entJob.GetAttributeValue<DateTime>("createdon");
                if (entJob.Attributes.Contains("msdyn_customerasset"))
                {
                    Entity entCustAsset = _service.Retrieve(msdyn_customerasset.EntityLogicalName, entJob.GetAttributeValue<EntityReference>("msdyn_customerasset").Id, new ColumnSet("hil_invoicedate"));
                    if (entCustAsset != null)
                    {
                        if (entCustAsset.Attributes.Contains("hil_invoicedate"))
                        {
                            _assetPurchaseDate = entCustAsset.GetAttributeValue<DateTime>("hil_invoicedate");
                        }
                    }

                    qrExp = new QueryExpression("msdyn_workorderproduct");
                    qrExp.ColumnSet = new ColumnSet("msdyn_workorderproductid");
                    qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                    qrExp.Criteria.AddCondition("msdyn_workorder", ConditionOperator.Equal, _jobGuId);
                    qrExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                    qrExp.Criteria.AddCondition("hil_markused", ConditionOperator.Equal, true);
                    entCol = _service.RetrieveMultiple(qrExp);
                    if (entCol.Entities.Count > 0) { SparePartUsed = true; }

                    qrExp = new QueryExpression("msdyn_workorderincident");
                    qrExp.ColumnSet = new ColumnSet("msdyn_workorderincidentid");
                    qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                    qrExp.Criteria.AddCondition("msdyn_workorder", ConditionOperator.Equal, _jobGuId);
                    qrExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                    qrExp.Criteria.AddCondition("hil_warrantystatus", ConditionOperator.Equal, 3); // Warranty Void
                    entCol = _service.RetrieveMultiple(qrExp);
                    if (entCol.Entities.Count == 0)
                    {
                        qrExp = new QueryExpression("hil_unitwarranty");
                        qrExp.ColumnSet = new ColumnSet("hil_warrantytemplate", "hil_warrantystartdate", "hil_warrantyenddate");
                        qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                        qrExp.Criteria.AddCondition("hil_customerasset", ConditionOperator.Equal, entJob.GetAttributeValue<EntityReference>("msdyn_customerasset").Id);
                        //qrExp.Criteria.AddCondition("hil_warrantystartdate", ConditionOperator.OnOrBefore, new DateTime(_jobCreatedOn.Year, _jobCreatedOn.Month, _jobCreatedOn.Day));
                        //qrExp.Criteria.AddCondition("hil_warrantyenddate", ConditionOperator.OnOrAfter, new DateTime(_jobCreatedOn.Year, _jobCreatedOn.Month, _jobCreatedOn.Day));
                        qrExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                        entCol = _service.RetrieveMultiple(qrExp);
                        if (entCol.Entities.Count > 0)
                        {
                            foreach (Entity Wt in entCol.Entities)
                            {
                                DateTime iValidTo = (DateTime)Wt["hil_warrantyenddate"];
                                DateTime iValidFrom = (DateTime)Wt["hil_warrantystartdate"];
                                if (_jobCreatedOn >= iValidFrom && _jobCreatedOn <= iValidTo)
                                {
                                    _jobWarrantyStatus = 1;
                                    _unitWarranty = Wt.ToEntityReference();
                                    _warrantyTemplateId = Wt.GetAttributeValue<EntityReference>("hil_warrantytemplate").Id;
                                    Entity _entTemp = _service.Retrieve(hil_warrantytemplate.EntityLogicalName, _warrantyTemplateId, new ColumnSet("hil_type"));
                                    if (_entTemp != null)
                                    {
                                        _warrantyTempType = _entTemp.GetAttributeValue<OptionSetValue>("hil_type").Value;
                                        if (_warrantyTempType == 1) { _jobWarrantySubStatus = 1; }
                                        else if (_warrantyTempType == 2) { _jobWarrantySubStatus = 2; }
                                        else if (_warrantyTempType == 7) { _jobWarrantySubStatus = 3; }
                                        else if (_warrantyTempType == 3) { _jobWarrantySubStatus = 4; }
                                    }
                                    _unitWarrStartDate = Wt.GetAttributeValue<DateTime>("hil_warrantystartdate");
                                    TimeSpan difference = (_jobCreatedOn - _unitWarrStartDate);
                                    _jobMonth = Math.Round((difference.Days * 1.0 / 30.42), 0);
                                    qrExp = new QueryExpression("hil_labor");
                                    qrExp.ColumnSet = new ColumnSet("hil_laborid", "hil_includedinwarranty", "hil_validtomonths", "hil_validfrommonths");
                                    qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                                    qrExp.Criteria.AddCondition("hil_warrantytemplateid", ConditionOperator.Equal, _warrantyTemplateId);
                                    qrExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                                    entCol = _service.RetrieveMultiple(qrExp);
                                    if (entCol.Entities.Count == 0) { laborInWarranty = true; }
                                    else
                                    {
                                        if (_jobMonth >= entCol.Entities[0].GetAttributeValue<int>("hil_validfrommonths") && _jobMonth <= entCol.Entities[0].GetAttributeValue<int>("hil_validtomonths"))
                                        {
                                            OptionSetValue _laborType = entCol.Entities[0].GetAttributeValue<OptionSetValue>("hil_includedinwarranty");
                                            laborInWarranty = _laborType.Value == 1 ? true : false;
                                        }
                                        else
                                        {
                                            OptionSetValue _laborType = entCol.Entities[0].GetAttributeValue<OptionSetValue>("hil_includedinwarranty");
                                            laborInWarranty = !(_laborType.Value == 1 ? true : false);
                                        }
                                    }
                                    break;
                                }
                            }
                        }
                    }
                    else
                    {
                        _jobWarrantyStatus = 3;
                    }
                }
            }

            Entity entJobUpdate = new Entity(msdyn_workorder.EntityLogicalName, _jobGuId);
            entJobUpdate["hil_warrantystatus"] = new OptionSetValue(_jobWarrantyStatus);
            if (_jobWarrantyStatus == 1)
            {
                entJobUpdate["hil_warrantysubstatus"] = new OptionSetValue(_jobWarrantySubStatus);
            }
            if (_unitWarranty != null)
            {
                entJobUpdate["hil_unitwarranty"] = _unitWarranty;
            }
            entJobUpdate["hil_laborinwarranty"] = laborInWarranty;
            if (_assetPurchaseDate.Year != 1900)
            {
                entJobUpdate["hil_purchasedate"] = _assetPurchaseDate;
            }
            entJobUpdate["hil_sparepartuse"] = SparePartUsed;
            if (_opChannelPartnerCategory != null)
            {
                entJobUpdate["hil_channelpartnercategory"] = _opChannelPartnerCategory;
            }
            _service.Update(entJobUpdate);
        }

        public void CreateSAWActivity(Guid JobId, IOrganizationService _service)
        {
            string _strCreatedOn = string.Empty;
            DateTime createdOn;
            EntityReference erCustomerAsset = null;
            msdyn_workorder _enJob = (msdyn_workorder)_service.Retrieve(msdyn_workorder.EntityLogicalName, JobId, new ColumnSet("msdyn_customerasset", "createdon", "hil_laborinwarranty"));
            bool _underReview = false;
            bool _laborinwarranty = false;

            if (_enJob != null)
            {
                if (_enJob.Attributes.Contains("hil_laborinwarranty"))
                {
                    _laborinwarranty = _enJob.GetAttributeValue<bool>("hil_laborinwarranty");
                }
                if (_enJob.Attributes.Contains("createdon"))
                {
                    createdOn = _enJob.GetAttributeValue<DateTime>("createdon").AddDays(-90);
                    _strCreatedOn = createdOn.Year.ToString() + "-" + createdOn.Month.ToString().PadLeft(2, '0') + "-" + createdOn.Day.ToString().PadLeft(2, '0');
                }
                if (_enJob.Attributes.Contains("msdyn_customerasset"))
                {
                    erCustomerAsset = _enJob.GetAttributeValue<EntityReference>("msdyn_customerasset");
                }
            }
            if (_laborinwarranty)
            {
                string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                  <entity name='msdyn_workorder'>
                    <attribute name='msdyn_name' />
                    <filter type='and'>
                      <condition attribute='msdyn_customerasset' operator='eq' value='" + erCustomerAsset.Id + @"' />
                      <condition attribute='hil_callsubtype' operator='eq' value='{E3129D79-3C0B-E911-A94E-000D3AF06CD4}' />
                      <condition attribute='msdyn_timeclosed' operator='on-or-after' value='" + _strCreatedOn + @"' />
                      <condition attribute='msdyn_substatus' operator='in'>
                        <value >{2927FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                        <value >{7E85074C-9C54-E911-A951-000D3AF0677F}</value>
                        <value >{1727FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                      </condition>
                    </filter>
                  </entity>
                </fetch>";

                EntityCollection entCol = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (entCol.Entities.Count > 0)
                {
                    string _remarks = "Old Job# " + entCol.Entities[0].GetAttributeValue<string>("msdyn_name");
                    CommonLib obj = new CommonLib();
                    CommonLib objReturn = obj.CreateSAWActivity(JobId, 0, SAWCategoryConst._GasChargePreAuditPastInstallationHistory, _service, _remarks, entCol.Entities[0].ToEntityReference());
                    if (objReturn.statusRemarks == "OK")
                    {
                        _underReview = true;
                    }
                }

                _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                  <entity name='msdyn_workorder'>
                    <attribute name='msdyn_name' />
                    <filter type='and'>
                      <condition attribute='msdyn_customerasset' operator='eq' value='" + erCustomerAsset.Id + @"' />
                      <condition attribute='msdyn_timeclosed' operator='on-or-after' value='" + _strCreatedOn + @"' />
                      <condition attribute='msdyn_workorderid' operator='ne' value='" + JobId + @"' />
                      <condition attribute='msdyn_substatus' operator='in'>
                        <value >{2927FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                        <value >{7E85074C-9C54-E911-A951-000D3AF0677F}</value>
                        <value >{1727FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                      </condition>
                      <condition attribute='hil_isgascharged' operator='eq' value='1' />
                    </filter>
                  </entity>
                </fetch>";

                entCol = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (entCol.Entities.Count > 0)
                {
                    string _remarks = "Old Job# " + entCol.Entities[0].GetAttributeValue<string>("msdyn_name");
                    CommonLib obj = new CommonLib();
                    CommonLib objReturn = obj.CreateSAWActivity(JobId, 0, SAWCategoryConst._GasChargePreAuditPastRepeatHistory, _service, _remarks, entCol.Entities[0].ToEntityReference());
                    if (objReturn.statusRemarks == "OK")
                    {
                        _underReview = true;
                    }
                }
                if (_underReview)
                {
                    Entity Ticket = new Entity("msdyn_workorder");
                    Ticket.Id = JobId;
                    Ticket["hil_claimstatus"] = new OptionSetValue(1); //Claim Under Review
                    _service.Update(Ticket);
                }
            }
        }

        public void UpdateJobclosureSourceforMobileApp(IOrganizationService service, Guid TicketId)
        {
            try
            {
                Entity Ticket = new Entity("msdyn_workorder");
                Ticket.Id = TicketId;
                QueryExpression Query;
                EntityCollection entcoll;

                Query = new QueryExpression("msdyn_workorder");
                Query.ColumnSet = new ColumnSet("entityimage");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition("msdyn_workorderid", ConditionOperator.Equal, Ticket.Id);
                entcoll = service.RetrieveMultiple(Query);
                foreach (Entity ent in entcoll.Entities)
                {
                    if (ent.Attributes.Contains("entityimage") || ent.Contains("entityimage"))
                    {
                        Ticket["hil_requesttype"] = new OptionSetValue(4);
                        service.Update(Ticket);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        public void UpdateJobUpcountry(IOrganizationService service, Guid TicketId)
        {
            try
            {
                Entity Ticket = new Entity("msdyn_workorder");
                Ticket.Id = TicketId;

                QueryExpression Query;
                EntityCollection entcoll;

                Query = new QueryExpression("msdyn_workorder");
                Query.ColumnSet = new ColumnSet("hil_productcategory", "hil_pincode", "hil_owneraccount", "hil_callsubtype", "hil_brand");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition("msdyn_workorderid", ConditionOperator.Equal, TicketId);
                entcoll = service.RetrieveMultiple(Query);
                foreach (Entity ent in entcoll.Entities)
                {
                    if (ent.GetAttributeValue<OptionSetValue>("hil_brand").Value == 2) //LLOYD
                    {
                        if (ent.Attributes.Contains("hil_productcategory") && ent.Attributes.Contains("hil_pincode") && ent.Attributes.Contains("hil_owneraccount") && ent.Attributes.Contains("hil_callsubtype"))
                        {
                            QueryExpression Query1;
                            EntityCollection entcoll1;
                            Query1 = new QueryExpression("hil_assignmentmatrix");
                            Query1.ColumnSet = new ColumnSet("hil_upcountry");
                            Query1.Criteria = new FilterExpression(LogicalOperator.And);
                            Query1.Criteria.AddCondition("hil_division", ConditionOperator.Equal, ent.GetAttributeValue<EntityReference>("hil_productcategory").Id);
                            Query1.Criteria.AddCondition("hil_pincode", ConditionOperator.Equal, ent.GetAttributeValue<EntityReference>("hil_pincode").Id);
                            Query1.Criteria.AddCondition("hil_franchiseedirectengineer", ConditionOperator.Equal, ent.GetAttributeValue<EntityReference>("hil_owneraccount").Id);
                            Query1.Criteria.AddCondition("hil_callsubtype", ConditionOperator.Equal, ent.GetAttributeValue<EntityReference>("hil_callsubtype").Id);
                            Query1.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                            entcoll1 = service.RetrieveMultiple(Query1);
                            if (entcoll1.Entities.Count > 0)
                            {
                                if (entcoll1.Entities[0].Attributes.Contains("hil_upcountry"))
                                {
                                    bool flag = entcoll1.Entities[0].GetAttributeValue<bool>("hil_upcountry");
                                    Ticket["hil_countryclassification"] = flag == true ? new OptionSetValue(2) : new OptionSetValue(1);
                                }
                                else
                                {
                                    Ticket["hil_countryclassification"] = new OptionSetValue(1); // Local
                                }
                            }
                            else
                            {
                                Ticket["hil_countryclassification"] = new OptionSetValue(2); // Upcountry
                            }
                            Ticket["hil_termsconditions"] = "Done";
                            service.Update(Ticket);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        public void UpdateClaimParams(IOrganizationService _service, Guid _jobGuId)
        {
            try
            {
                bool SparePartUsed = false;
                DateTime _fromDate, _toDate;
                QueryExpression qrExp;
                EntityCollection entCol;
                int ClosureCount = 0;

                Entity entJob = _service.Retrieve(msdyn_workorder.EntityLogicalName, _jobGuId, new ColumnSet("msdyn_timeclosed", "createdon", "hil_owneraccount"));
                if (entJob != null)
                {
                    DateTime _jobCreatedOn = entJob.GetAttributeValue<DateTime>("createdon");
                    DateTime _jobClosedOn = entJob.GetAttributeValue<DateTime>("msdyn_timeclosed");

                    if (DateTime.Now.Day >= 21)
                    {
                        _fromDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 21);
                        _toDate = _fromDate.AddDays(30);
                    }
                    else
                    {
                        _toDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 20);
                        _fromDate = _toDate.AddDays(-30);
                    }

                    qrExp = new QueryExpression("hil_claimperiod");
                    qrExp.ColumnSet = new ColumnSet("hil_fromdate", "hil_todate");
                    qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                    qrExp.Criteria.AddCondition("hil_fromdate", ConditionOperator.OnOrBefore, _jobCreatedOn);
                    qrExp.Criteria.AddCondition("hil_todate", ConditionOperator.OnOrAfter, _jobCreatedOn);
                    entCol = _service.RetrieveMultiple(qrExp);
                    if (entCol.Entities.Count > 0)
                    {
                        _fromDate = entCol.Entities[0].GetAttributeValue<DateTime>("hil_fromdate");
                        _toDate = entCol.Entities[0].GetAttributeValue<DateTime>("hil_todate").AddDays(1).AddMinutes(-1);
                    }

                    qrExp = new QueryExpression("msdyn_workorderproduct");
                    qrExp.ColumnSet = new ColumnSet("msdyn_workorderproductid");
                    qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                    qrExp.Criteria.AddCondition("msdyn_workorder", ConditionOperator.Equal, _jobGuId);
                    qrExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                    qrExp.Criteria.AddCondition("hil_markused", ConditionOperator.Equal, true);
                    entCol = _service.RetrieveMultiple(qrExp);
                    if (entCol.Entities.Count > 0) { SparePartUsed = true; }

                    qrExp = new QueryExpression("msdyn_workorder");
                    qrExp.ColumnSet = new ColumnSet("msdyn_workorderid", "msdyn_timeclosed");
                    qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                    qrExp.Criteria.AddCondition("msdyn_substatus", ConditionOperator.Equal, new Guid("1727FA6C-FA0F-E911-A94E-000D3AF060A1"));
                    qrExp.Criteria.AddCondition("hil_owneraccount", ConditionOperator.Equal, entJob.GetAttributeValue<EntityReference>("hil_owneraccount").Id);
                    qrExp.Criteria.AddCondition("msdyn_timeclosed", ConditionOperator.OnOrAfter, _fromDate);
                    qrExp.Criteria.AddCondition("msdyn_timeclosed", ConditionOperator.OnOrBefore, _toDate);
                    qrExp.Criteria.AddCondition("msdyn_timeclosed", ConditionOperator.OnOrBefore, _jobClosedOn);
                    qrExp.AddOrder("msdyn_timeclosed", OrderType.Ascending);
                    entCol = _service.RetrieveMultiple(qrExp);
                    if (entCol.Entities.Count > 0)
                    {
                        foreach (Entity ent in entCol.Entities)
                        {
                            if (ent.GetAttributeValue<DateTime>("msdyn_timeclosed").AddMinutes(330) <= _jobClosedOn)
                            {
                                ClosureCount++;
                            }
                        }
                    }

                    Entity _jobEnt = new Entity(msdyn_workorder.EntityLogicalName, _jobGuId);
                    _jobEnt["hil_sparepartuse"] = SparePartUsed;
                    _jobEnt["hil_jobclosurecounter"] = ClosureCount;
                    _service.Update(_jobEnt);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        #region De Link PO
        public static void DeLinkAssociatedPO(IOrganizationService service, Entity entity)
        {
            bool IfDeLink = false;
            msdyn_workorder iWo = entity.ToEntity<msdyn_workorder>();
            hil_productrequestheader iHeader = new hil_productrequestheader();
            if (entity.Contains("hil_delinkpo") && entity.Attributes.Contains("hil_delinkpo"))
            {
                IfDeLink = (bool)entity.GetAttributeValue<bool>("hil_delinkpo");
                if (IfDeLink == true)
                {
                    #region To restrict user to delink Replacement Approved PO
                    //========Start=========
                    /* 
                     * Date:17/Oct/2019
                     * UserId: Kuldeep Khare
                     * Remarks: To restrict user to delink Replacement Approved PO
                     */
                    QueryExpression QueryPR = new QueryExpression(hil_productrequestheader.EntityLogicalName);
                    QueryPR.ColumnSet = new ColumnSet("hil_job", "hil_prtype", "statuscode");
                    QueryPR.Criteria = new FilterExpression(LogicalOperator.And);
                    QueryPR.Criteria.AddCondition(new ConditionExpression("hil_job", ConditionOperator.Equal, entity.Id));
                    QueryPR.Criteria.AddCondition(new ConditionExpression("hil_prtype", ConditionOperator.Equal, 910590002));
                    EntityCollection FoundPR = service.RetrieveMultiple(QueryPR);

                    if (FoundPR.Entities.Count > 0)
                    {
                        foreach (hil_productrequestheader prHeader in FoundPR.Entities)
                        {
                            if (prHeader.Contains("statuscode") && prHeader.statuscode.Value == 910590003)
                            {
                                throw new InvalidPluginExecutionException("  ***Respective PO cannot be delinked as replacement is already approved***  ");
                            }
                        }
                    }
                    //======== End ===========
                    #endregion  

                    QueryExpression Query = new QueryExpression(hil_productrequest.EntityLogicalName);
                    Query.ColumnSet = new ColumnSet("hil_job", "hil_prheader");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition(new ConditionExpression("hil_job", ConditionOperator.Equal, entity.Id));
                    Query.Criteria.AddCondition(new ConditionExpression("hil_prheader", ConditionOperator.NotNull));
                    EntityCollection Found = service.RetrieveMultiple(Query);
                    if (Found.Entities.Count > 0)
                    {
                        foreach (hil_productrequest iReq in Found.Entities)
                        {
                            iReq.hil_Job = null;
                            iHeader.Id = iReq.hil_PRHeader.Id;
                            iHeader.hil_Job = null;
                            service.Update(iHeader);
                            service.Update(iReq);
                        }
                        Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Work Initiated", entity.Id, service);
                        iWo.hil_FlagPO = false;
                        service.Update(iWo);
                    }
                    else
                    {
                        Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Work Initiated", entity.Id, service);
                        iWo.hil_FlagPO = false;
                        service.Update(iWo);
                    }
                }
            }
        }
        #endregion
        #region Call Allocation
        //public static void CallAllocation(IOrganizationService service, Entity entity)
        //{
        //    //1 - Both
        //    //2 - Direct Engineer
        //    //3 - Franchisee
        //    if (entity.Attributes.Contains("hil_automaticassign"))
        //    {
        //        Guid iPin = new Guid();
        //        Guid iDivision = new Guid();
        //        Guid iCallStype = new Guid();
        //        EntityReference iAssign = new EntityReference(SystemUser.EntityLogicalName);
        //        int iResource = new int();
        //        OptionSetValue IfAssign = (OptionSetValue)entity["hil_automaticassign"];
        //        if (IfAssign.Value == 1)
        //        {
        //            msdyn_workorder iJob = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, entity.Id, new ColumnSet(true));
        //            if (iJob.hil_pincode != null)
        //                iPin = iJob.hil_pincode.Id;
        //            if (iJob.hil_Productcategory != null)
        //                iDivision = iJob.hil_Productcategory.Id;
        //            if (iJob.hil_CallSubType != null)
        //                iCallStype = iJob.hil_CallSubType.Id;
        //            if (iDivision != Guid.Empty)
        //            {
        //                Product pdt = (Product)service.Retrieve(Product.EntityLogicalName, iDivision, new ColumnSet("hil_repairresource"));
        //                if (pdt.hil_repairresource != null)
        //                    iResource = pdt.hil_repairresource.Value;
        //                switch (iResource)
        //                {
        //                    case 1: //Both
        //                        iAssign = GetAssigneeForBoth(service, iPin, iDivision, iCallStype, iJob.Id, iJob.OwnerId);
        //                        break;
        //                    case 2: //Direct Engineer
        //                        iAssign = GetAssignee(service, iPin, iDivision, iCallStype, iResource, iJob.Id, iJob.OwnerId);
        //                        break;
        //                    case 3: //Franchisee
        //                        iAssign = GetAssignee(service, iPin, iDivision, iCallStype, iResource, iJob.Id, iJob.OwnerId);
        //                        break;
        //                }
        //                if(iAssign.Id != Guid.Empty)
        //                {
        //                    Helper.Assign(iAssign.LogicalName, iJob.LogicalName, iAssign.Id, iJob.Id, service);
        //                }
        //                else
        //                {
        //                    Helper.Assign(iAssign.LogicalName, iJob.LogicalName, new Guid("A188F3F3-368B-E811-A95A-000D3AF069BD"), iJob.Id, service);
        //                }
        //            }
        //        }
        //    }
        //}
        //public static EntityReference GetAssigneeForBoth(IOrganizationService service, Guid iPin, Guid iDivision, Guid iCallStype, Guid iJob, EntityReference Default)
        //{
        //    EntityReference Assignee = new EntityReference();
        //    msdyn_workorder eJob = new msdyn_workorder();
        //    eJob.Id = iJob;
        //    hil_assignmentmatrix iMatx = new hil_assignmentmatrix();
        //    QueryExpression Query = new QueryExpression(hil_assignmentmatrix.EntityLogicalName);
        //    Query.ColumnSet = new ColumnSet("hil_franchiseedirectengineer", "ownerid");
        //    Query.Criteria = new FilterExpression(LogicalOperator.And);
        //    Query.Criteria.AddCondition(new ConditionExpression("hil_pincode", ConditionOperator.Equal, iPin));
        //    Query.Criteria.AddCondition(new ConditionExpression("hil_division", ConditionOperator.Equal, iDivision));
        //    Query.Criteria.AddCondition(new ConditionExpression("hil_callsubtype", ConditionOperator.Equal, iCallStype));
        //    EntityCollection Found = service.RetrieveMultiple(Query);
        //    if (Found.Entities.Count == 1)
        //    {
        //        iMatx = Found.Entities[0].ToEntity<hil_assignmentmatrix>();
        //        if (iMatx.Attributes.Contains("hil_franchiseedirectengineer"))
        //        {
        //            EntityReference iRef = (EntityReference)iMatx["hil_franchiseedirectengineer"];
        //            Account iFown = (Account)service.Retrieve(Account.EntityLogicalName, iRef.Id, new ColumnSet("ownerid", "customertypecode"));
        //            if(iFown.CustomerTypeCode.Value == 9)
        //            {
        //                bool IfDEPresent = CheckIfDirectEngineerPresent(service, iFown.OwnerId.Id);
        //                if (IfDEPresent)
        //                {
        //                    Assignee = iFown.OwnerId;
        //                    eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx.Id);
        //                    service.Update(eJob);
        //                }
        //                else
        //                {
        //                    Assignee = iMatx.OwnerId;
        //                    eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx.Id);
        //                    service.Update(eJob);
        //                }
        //            }
        //            else
        //            {
        //                Assignee = iFown.OwnerId;
        //                eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx.Id);
        //            }
        //        }
        //        else
        //        {
        //            Assignee = iMatx.OwnerId;
        //        }
        //    }
        //    else if (Found.Entities.Count > 1)
        //    {
        //        iMatx = Found.Entities[0].ToEntity<hil_assignmentmatrix>();
        //        Assignee = iMatx.OwnerId;
        //        eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx.Id);
        //    }
        //    else if (Found.Entities.Count == 0)
        //    {
        //        QueryExpression Query1 = new QueryExpression(hil_assignmentmatrix.EntityLogicalName);
        //        Query1.ColumnSet = new ColumnSet("hil_franchiseedirectengineer", "ownerid");
        //        Query1.Criteria = new FilterExpression(LogicalOperator.And);
        //        Query1.Criteria.AddCondition(new ConditionExpression("hil_pincode", ConditionOperator.Equal, iPin));
        //        Query1.Criteria.AddCondition(new ConditionExpression("hil_division", ConditionOperator.Equal, iDivision));
        //        EntityCollection Found1 = service.RetrieveMultiple(Query1);
        //        if (Found1.Entities.Count > 0)
        //        {
        //            iMatx = Found1.Entities[0].ToEntity<hil_assignmentmatrix>();
        //            Assignee = iMatx.OwnerId;
        //            eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx.Id);
        //        }
        //        else
        //        {
        //            Assignee = Default;
        //        }
        //    }
        //    service.Update(eJob);
        //    return Assignee;
        //}
        //public static EntityReference GetAssignee(IOrganizationService service, Guid iPin, Guid iDivision, Guid iCallStype, int iResource, Guid iJob, EntityReference Default)
        //{
        //    EntityReference Assignee = new EntityReference();
        //    msdyn_workorder eJob = new msdyn_workorder();
        //    eJob.Id = iJob;
        //    hil_assignmentmatrix iMatx = new hil_assignmentmatrix();
        //    QueryExpression Query = new QueryExpression(hil_assignmentmatrix.EntityLogicalName);
        //    Query.ColumnSet = new ColumnSet("hil_franchiseedirectengineer", "ownerid");
        //    Query.Criteria = new FilterExpression(LogicalOperator.And);
        //    Query.Criteria.AddCondition(new ConditionExpression("hil_pincode", ConditionOperator.Equal, iPin));
        //    Query.Criteria.AddCondition(new ConditionExpression("hil_division", ConditionOperator.Equal, iDivision));
        //    Query.Criteria.AddCondition(new ConditionExpression("hil_callsubtype", ConditionOperator.Equal, iCallStype));
        //    EntityCollection Found = service.RetrieveMultiple(Query);
        //    if (Found.Entities.Count > 1)
        //    {
        //        foreach (hil_assignmentmatrix iMatx1 in Found.Entities)
        //        {
        //            if (iMatx1.Attributes.Contains("hil_franchiseedirectengineer"))
        //            {
        //                EntityReference iRef = (EntityReference)iMatx1["hil_franchiseedirectengineer"];
        //                Account iFown = (Account)service.Retrieve(Account.EntityLogicalName, iRef.Id, new ColumnSet("ownerid"));
        //                SystemUser iSys = (SystemUser)service.Retrieve(SystemUser.EntityLogicalName, iFown.OwnerId.Id, new ColumnSet("positionid"));
        //                if (iSys.PositionId != null)
        //                {
        //                    if ((iResource == 2) && (iSys.PositionId.Name == "Direct Engineer"))
        //                    {
        //                        bool IfDEPresent = CheckIfDirectEngineerPresent(service, iFown.OwnerId.Id);
        //                        if(IfDEPresent)
        //                        {
        //                            Assignee = iFown.OwnerId;
        //                            eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx1.Id);
        //                            service.Update(eJob);
        //                            break;
        //                        }
        //                        else
        //                        {
        //                            Assignee = iMatx1.OwnerId;
        //                            eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx1.Id);
        //                            service.Update(eJob);
        //                            break;
        //                        }
        //                    }
        //                    else if ((iResource == 3) && (iSys.PositionId.Name == "Franchisee"))
        //                    {
        //                        Assignee = iFown.OwnerId;
        //                        eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx1.Id);
        //                        service.Update(eJob);
        //                        break;
        //                    }
        //                    else
        //                    {
        //                        Assignee = iMatx1.OwnerId;  //Branch Head
        //                        eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx1.Id);
        //                        service.Update(eJob);
        //                    }
        //                }
        //                else
        //                {
        //                    Assignee = iMatx1.OwnerId;  //Branch Head
        //                    eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx1.Id);
        //                    service.Update(eJob);
        //                }
        //            }
        //            else
        //            {
        //                Assignee = iMatx1.OwnerId;  //Branch Head
        //                eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx1.Id);
        //                service.Update(eJob);
        //            }
        //        }
        //    }
        //    else if (Found.Entities.Count == 1)
        //    {
        //        iMatx = Found.Entities[0].ToEntity<hil_assignmentmatrix>();
        //        if (iMatx.Attributes.Contains("hil_franchiseedirectengineer"))
        //        {
        //            EntityReference iRef = (EntityReference)iMatx["hil_franchiseedirectengineer"];
        //            if(iRef != null)
        //            {
        //                Account iFown = (Account)service.Retrieve(Account.EntityLogicalName, iRef.Id, new ColumnSet("ownerid"));
        //                SystemUser iSys = (SystemUser)service.Retrieve(SystemUser.EntityLogicalName, iFown.OwnerId.Id, new ColumnSet("positionid"));
        //                if (iSys.PositionId != null)
        //                {
        //                    if ((iResource == 2) && (iSys.PositionId.Name == "Direct Engineer"))
        //                    {
        //                        Assignee = iFown.OwnerId;
        //                        eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx.Id);
        //                        service.Update(eJob);
        //                    }
        //                    else if ((iResource == 3) && (iSys.PositionId.Name == "Franchisee"))
        //                    {
        //                        Assignee = iFown.OwnerId;
        //                        eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx.Id);
        //                        service.Update(eJob);
        //                    }
        //                    else
        //                    {
        //                        Assignee = iMatx.OwnerId; //Branch Head
        //                        eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx.Id);
        //                        service.Update(eJob);
        //                    }
        //                }
        //                else
        //                {
        //                    Assignee = iMatx.OwnerId; //Branch Head
        //                    eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx.Id);
        //                    service.Update(eJob);
        //                }
        //            }
        //            else
        //            {
        //                Assignee = iMatx.OwnerId; //Branch Head
        //                eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx.Id);
        //                service.Update(eJob);
        //            }
        //        }
        //        else
        //        {
        //            Assignee = iMatx.OwnerId; //Branch Head
        //            eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx.Id);
        //            service.Update(eJob);
        //        }
        //    }
        //    else if (Found.Entities.Count == 0)
        //    {
        //        QueryExpression Query1 = new QueryExpression(hil_assignmentmatrix.EntityLogicalName);
        //        Query1.ColumnSet = new ColumnSet("hil_franchiseedirectengineer", "ownerid");
        //        Query1.Criteria = new FilterExpression(LogicalOperator.And);
        //        Query1.Criteria.AddCondition(new ConditionExpression("hil_pincode", ConditionOperator.Equal, iPin));
        //        Query1.Criteria.AddCondition(new ConditionExpression("hil_division", ConditionOperator.Equal, iDivision));
        //        EntityCollection Found1 = service.RetrieveMultiple(Query1);
        //        if (Found1.Entities.Count > 0)
        //        {
        //            iMatx = Found1.Entities[0].ToEntity<hil_assignmentmatrix>();
        //            Assignee = iMatx.OwnerId;
        //            eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx.Id);
        //            service.Update(eJob);
        //        }
        //        else
        //        {
        //            Assignee = Default;
        //        }
        //     }
        //    return Assignee;
        //}
        #endregion
        #region WorkOrderPO
        public static void WorkOrderPO(Entity entity, IOrganizationService service)
        {
            try
            {
                //tracingService.Trace("3");
                msdyn_workorder enWorkorder = entity.ToEntity<msdyn_workorder>();
                if (enWorkorder.hil_FlagPO == null)
                    return;
                if (enWorkorder.hil_FlagPO.Value == true)
                {
                    using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                    {
                        var obj = from _woProduct in orgContext.CreateQuery<msdyn_workorderproduct>()
                                  join _WorkOrder in orgContext.CreateQuery<msdyn_workorder>() on _woProduct.msdyn_WorkOrder.Id equals _WorkOrder.Id
                                  join _Account in orgContext.CreateQuery<Account>() on _WorkOrder.hil_OwnerAccount.Id equals _Account.Id
                                  join _Owner in orgContext.CreateQuery<SystemUser>() on _Account.OwnerId.Id equals _Owner.Id
                                  where _woProduct.msdyn_WorkOrder.Id == entity.Id
                                  && _woProduct.hil_purchaseorder == null
                                  && _woProduct.hil_AvailabilityStatus.Value == 2
                                  select new
                                  {
                                      _woProduct.msdyn_Product,
                                      _woProduct.msdyn_workorderproductId,
                                      _WorkOrder.OwnerId,
                                      _Owner.SystemUserId,
                                      _Account.AccountId,
                                      _woProduct.msdyn_Quantity,
                                      _woProduct.hil_WarrantyStatus,
                                      _Account.CustomerTypeCode,
                                      _Account.hil_InWarrantyCustomerSAPCode,
                                      _Account.hil_OutWarrantyCustomerSAPCode
                                  };
                        //tracingService.Trace("4");
                        EntityReference _productRequestHeader = null;
                        bool _prExist = false;
                        foreach (var iobj in obj)
                        {
                            string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='hil_productrequest'>
                            <attribute name='hil_prheader' />
                            <filter type='and'>
                                <condition attribute='hil_job' operator='eq' value='{entity.Id}' />
                                <condition attribute='hil_partcode' operator='eq' value='{iobj.msdyn_Product.Id}' />
                                <condition attribute='statecode' operator='eq' value='0' />
                            </filter>
                            </entity>
                            </fetch>";
                            EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                            if (entCol.Entities.Count > 0) { 
                            
                            }
                        }
                        foreach (var iobj in obj)
                        {
                            #region Variables
                            Guid fsPart = Guid.Empty;
                            Guid fsOwnerAccount = Guid.Empty;
                            Guid fsOwner = Guid.Empty;
                            OptionSetValue opWarrantyStatus = null;
                            OptionSetValue opAccountType = null; // CUstomer type code of franchisee
                            OptionSetValue opTecDEAccounType = null;// CUstomer type code of technician or direct engineer
                            Int32 iQuantityRequired = 0;
                            String sCustomerSAPCOde = String.Empty;
                            String sDistributionChannel = String.Empty;

                            #endregion

                            if (iobj.msdyn_Product != null)
                            {
                                fsPart = iobj.msdyn_Product.Id;
                            }
                            if (iobj.SystemUserId != null)
                            {
                                fsOwner = iobj.SystemUserId.Value;
                            }
                            if (iobj.AccountId != null)
                            {
                                fsOwnerAccount = iobj.AccountId.Value;
                            }
                            if (iobj.msdyn_Quantity != null)
                            {
                                iQuantityRequired = (Int32)iobj.msdyn_Quantity.Value;
                            }
                            if (iobj.hil_WarrantyStatus != null)
                            {
                                opWarrantyStatus = iobj.hil_WarrantyStatus;
                                if (opWarrantyStatus.Value == 1)
                                {
                                    if (iobj.hil_InWarrantyCustomerSAPCode != null)
                                    {
                                        sCustomerSAPCOde = iobj.hil_InWarrantyCustomerSAPCode;
                                    }
                                    else
                                    {
                                        throw new InvalidPluginExecutionException("XXXXXXXXXXXXXXXXXXXXXXXX - FRANCHISE SAP CODE BLANK IN MASTER PLEASE CONTACT SYSTEM ADMINISTRATOR - XXXXXXXXXXXXXXXXXXXXXXXX");
                                    }
                                }
                                else if (opWarrantyStatus.Value == 2)
                                {
                                    //Added on 25th February 2019 for "ALL OUTGOING POs SHOULD BE IN WARRANTY"
                                    if (iobj.hil_InWarrantyCustomerSAPCode != null)
                                    {
                                        sCustomerSAPCOde = iobj.hil_InWarrantyCustomerSAPCode;
                                    }
                                    else
                                    {
                                        throw new InvalidPluginExecutionException("XXXXXXXXXXXXXXXXXXXXXXXX - FRANCHISE SAP CODE BLANK IN MASTER PLEASE CONTACT SYSTEM ADMINISTRATOR - XXXXXXXXXXXXXXXXXXXXXXXX");
                                    }
                                    //sCustomerSAPCOde = iobj.hil_OutWarrantyCustomerSAPCode;
                                }
                                //Added on 25th February 2019 for "ALL OUTGOING POs SHOULD BE IN WARRANTY"
                                opWarrantyStatus = new OptionSetValue(1);
                            }
                            if (iobj.CustomerTypeCode != null)
                            {
                                opAccountType = iobj.CustomerTypeCode;
                            }
                            sDistributionChannel = HelperPO.getDistributionChannel(opWarrantyStatus, opAccountType, service);
                            opTecDEAccounType = Helper.GetAccountType(iobj.OwnerId.Id, service);
                            if (opTecDEAccounType != null)
                            {
                                if (opTecDEAccounType.Value != 5 && opTecDEAccounType.Value != 6 && opTecDEAccounType.Value != 9)
                                {
                                    throw new InvalidPluginExecutionException("Channel Partner Master Setup Error !!! " + iobj.OwnerId.Name + " is not mapped as ASP/DSE");
                                }
                            }
                            else {
                                throw new InvalidPluginExecutionException("Channel Partner Master Setup Error !!! " + iobj.OwnerId.Name + " is not mapped as ASP/DSE");
                            }
                            Guid fsPOId = HelperPO.CreatePO(service, fsOwner, entity.Id, Guid.Empty, iQuantityRequired, fsOwnerAccount, fsPart, 910590003, opTecDEAccounType, sCustomerSAPCOde, sDistributionChannel, opWarrantyStatus);
                            
                            Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Part PO Created", entity.Id, service);
                            #region UpdatePOonWOProduct
                            msdyn_workorderproduct upWOProduct = new msdyn_workorderproduct();
                            upWOProduct.Id = iobj.msdyn_workorderproductId.Value;
                            upWOProduct.hil_purchaseorder = new EntityReference(hil_productrequest.EntityLogicalName, fsPOId);
                            service.Update(upWOProduct);
                            //tracingService.Trace("5");
                            #endregion
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }

        }
        #endregion
        #region Work Order AMC
        //public static void WorkOrderAMC(Entity entity, IOrganizationService service)
        //{
        //    msdyn_workorder enWorkorder = entity.ToEntity<msdyn_workorder>();
        //    if (entity.Attributes.Contains("hil_repairdone"))
        //    {
        //        if (enWorkorder.hil_RepairDone.Value == 1)
        //        {
        //            LookForWorkOrderIncidents(service, enWorkorder);
        //        }
        //    }
        //}
        //public static void LookForWorkOrderIncidents(IOrganizationService service, msdyn_workorder _Workorder)
        //{
        //    bool _IfAMCEligible = false;
        //    using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
        //    {
        //        Guid fsResult = Guid.Empty;
        //        var obj = from _WO in orgContext.CreateQuery<msdyn_workorder>()
        //                  join _WoIn in orgContext.CreateQuery<msdyn_workorderincident>() on _WO.Id equals _WoIn.msdyn_WorkOrder.Id
        //                  join _CustAst in orgContext.CreateQuery<msdyn_customerasset>() on _WoIn.msdyn_CustomerAsset.Id equals _CustAst.Id
        //                  join _UwRty in orgContext.CreateQuery<hil_unitwarranty>() on _CustAst.Id equals _UwRty.hil_CustomerAsset.Id
        //                  join _WrtyTmplt in orgContext.CreateQuery<hil_warrantytemplate>() on _UwRty.hil_WarrantyTemplate.Id equals _WrtyTmplt.Id
        //                  where _UwRty.statuscode.Value == 1 //In Warranty
        //                  && _UwRty.hil_ProductType.Value == 1 //Finished Good
        //                  && _WrtyTmplt.hil_type.Value != 3 //Not AMC
        //                  select new
        //                  {
        //                      _CustAst.hil_InvoiceDate,
        //                      _WoIn.msdyn_CustomerAsset,
        //                      _CustAst.hil_ProductCategory,
        //                      _CustAst.hil_Customer
        //                  };
        //        foreach (var iobj in obj)
        //        {
        //            if (iobj.hil_InvoiceDate != null && iobj.msdyn_CustomerAsset != null && iobj.hil_ProductCategory != null)
        //            {
        //                _IfAMCEligible = EvaluateAMC(service, Convert.ToDateTime(iobj.hil_InvoiceDate), iobj.msdyn_CustomerAsset, iobj.hil_ProductCategory, iobj.hil_Customer);
        //            }
        //        }
        //    }
        //    if (_IfAMCEligible)
        //        _Workorder.hil_IfAMCEligible = new OptionSetValue(1);
        //    else
        //        _Workorder.hil_IfAMCEligible = new OptionSetValue(2);
        //    service.Update(_Workorder);
        //}
        //public static bool EvaluateAMC(IOrganizationService service, DateTime InvDate, EntityReference CustAsst, EntityReference PdtCgry, EntityReference Customer)
        //{
        //    bool IsEligible = false;
        //    IsEligible = CheckNumberOfYearsForAMC(service, InvDate, PdtCgry, Customer);
        //    if (IsEligible == true)
        //    {
        //        IsEligible = IfNumberOfRepairsExceeds(service, InvDate, PdtCgry, Customer);
        //    }
        //    return (IsEligible);
        //}
        //public static bool CheckNumberOfYearsForAMC(IOrganizationService service, DateTime InvDate, EntityReference PdtCtgry, EntityReference Customer)
        //{
        //    bool IsEligible = false;
        //    Product _ThisPdCgry = (Product)service.Retrieve(Product.EntityLogicalName, PdtCtgry.Id, new ColumnSet("hil_yearsforwarranty"));
        //    decimal YrsOfWrty = (decimal)_ThisPdCgry.hil_YearsforWarranty;
        //    int _NumMonths = Convert.ToInt32(YrsOfWrty * 12);
        //    DateTime CheckDate = InvDate.AddMonths(_NumMonths);
        //    if (CheckDate >= DateTime.Now)
        //    {
        //        IsEligible = true;
        //    }
        //    return (IsEligible);
        //}
        //public static bool IfNumberOfRepairsExceeds(IOrganizationService service, DateTime Invdate, EntityReference PdtCtgry, EntityReference Customer)
        //{
        //    bool IsEligible = false;
        //    int Counter = 0;
        //    msdyn_workorderincident Inc = new msdyn_workorderincident();
        //    using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
        //    {
        //        var obj = from _WoIn in orgContext.CreateQuery<msdyn_workorderincident>()
        //                  join _CustAst in orgContext.CreateQuery<msdyn_customerasset>() on _WoIn.msdyn_CustomerAsset.Id equals _CustAst.Id
        //                  where _CustAst.hil_Customer.Id == Customer.Id //Check for this specefic customer
        //                  select new
        //                  {
        //                      _CustAst.hil_Customer
        //                  };
        //        foreach (var iobj in obj)
        //        {
        //            Counter += 1;
        //        }
        //    }
        //    Product _ThisPdCgry = (Product)service.Retrieve(Product.EntityLogicalName, PdtCtgry.Id, new ColumnSet("hil_numberinrepairs"));
        //    int _NumAMC = (int)_ThisPdCgry.hil_NumberinRepairs;
        //    if (Counter < _NumAMC)
        //    {
        //        IsEligible = true;
        //    }
        //    return (IsEligible);
        //}
        #endregion
        #region Create TAT Record
        public static void CreateTatRecord(Entity entity, IPluginExecutionContext context, IOrganizationService service)
        {
            try
            {
                tracingService.Trace("3");
                using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                {
                    tracingService.Trace("4");
                    msdyn_workorder enWorkorder = entity.ToEntity<msdyn_workorder>();
                    tracingService.Trace("5");
                    if (enWorkorder.msdyn_SubStatus != null)
                    {
                        tracingService.Trace("6");
                        var obj = (from _JobTat in orgContext.CreateQuery<hil_jobtat>()
                                   where _JobTat.hil_WorkOrder.Id == entity.Id
                                   orderby _JobTat.CreatedOn descending
                                   select new
                                   {
                                       _JobTat.hil_ToStatus
                                       ,
                                       _JobTat.hil_ToStatusTime
                                   }).Take(1);
                        tracingService.Trace("7" + Enumerable.Count(obj));
                        if (Enumerable.Count(obj) == 0)
                        {
                            hil_jobtat crJobTat = new hil_jobtat();
                            crJobTat.hil_ToStatusTime = DateTime.Now;
                            crJobTat.hil_ToStatus = enWorkorder.msdyn_SubStatus;
                            crJobTat.hil_WorkOrder = new EntityReference(entity.LogicalName, entity.Id);
                            service.Create(crJobTat);
                        }
                        else
                        {
                            tracingService.Trace("8");
                            foreach (var iobj in obj)
                            {
                                tracingService.Trace("9");
                                if (iobj.hil_ToStatusTime != null && iobj.hil_ToStatus != null)
                                {
                                    tracingService.Trace("10");
                                    hil_jobtat crJobTat = new hil_jobtat();
                                    crJobTat.hil_FromStatusTime = iobj.hil_ToStatusTime;
                                    crJobTat.hil_FromStatus = iobj.hil_ToStatus;
                                    crJobTat.hil_ToStatusTime = DateTime.Now;
                                    crJobTat.hil_ToStatus = enWorkorder.msdyn_SubStatus;
                                    crJobTat.hil_WorkOrder = new EntityReference(entity.LogicalName, entity.Id);
                                    service.Create(crJobTat);
                                    tracingService.Trace("11");
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        #endregion
        #region Execute Call Allocation
        //public static void ExecuteCallAllocation(IOrganizationService service, Entity serviceTicket)
        //{
        //    msdyn_workorder ServiceTicket = serviceTicket.ToEntity<msdyn_workorder>();//(msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, serviceTicket.Id, new ColumnSet(true));
        //    EntityReference cust = new EntityReference();
        //    EntityReference callSubType = new EntityReference(hil_callsubtype.EntityLogicalName);
        //    EntityReference pinCode = new EntityReference(hil_pincode.EntityLogicalName);
        //    EntityReference Division = new EntityReference(Product.EntityLogicalName);
        //    Guid CharactersticId = new Guid();
        //    int repairResource = 0;
        //    if (ServiceTicket.hil_AutomaticAssign != null)
        //    {
        //        if (ServiceTicket.hil_AutomaticAssign.Value == 1)
        //        {
        //            ServiceTicket = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, serviceTicket.Id, new ColumnSet("hil_callsubtype", "hil_pincode", "hil_productcategory"));
        //            if (ServiceTicket.hil_CallSubType != null)
        //            {
        //                callSubType = ServiceTicket.hil_CallSubType;
        //            }
        //            if (ServiceTicket.hil_pincode != null)
        //            {
        //                pinCode = ServiceTicket.hil_pincode;
        //            }
        //            if (ServiceTicket.hil_Productcategory != null)
        //            {
        //                Division = ServiceTicket.hil_Productcategory;
        //            }
        //            Product pdt = (Product)service.Retrieve(Product.EntityLogicalName, Division.Id, new ColumnSet("hil_repairresource"));
        //            if (pdt.hil_repairresource != null)
        //            {
        //                repairResource = pdt.hil_repairresource.Value;
        //            }
        //            QueryExpression query = new QueryExpression();
        //            query.EntityName = hil_assignmentmatrix.EntityLogicalName;
        //            ColumnSet Col = new ColumnSet(true);
        //            query.ColumnSet = Col;
        //            query.Criteria = new FilterExpression(LogicalOperator.And);
        //            query.Criteria.AddCondition("hil_division", ConditionOperator.Equal, Division.Id);
        //            query.Criteria.AddCondition("hil_pincode", ConditionOperator.Equal, pinCode.Id);
        //            query.Criteria.AddCondition("hil_assignto", ConditionOperator.NotNull);
        //            EntityCollection results = service.RetrieveMultiple(query);
        //            if (results.Entities.Count >= 1)
        //            {
        //                foreach (hil_assignmentmatrix assgnmntMatrix in results.Entities)
        //                {
        //                    ServiceTicket.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, assgnmntMatrix.Id);
        //                    service.Update(ServiceTicket);
        //                    EntityReference assignTo = assgnmntMatrix.hil_AssignTo;
        //                    QueryExpression query1 = new QueryExpression();
        //                    query1.EntityName = Characteristic.EntityLogicalName;
        //                    query1.ColumnSet = new ColumnSet("hil_assignto", "hil_productdivision", "hil_calltype", "hil_franchiseedirectengineer");
        //                    query1.Criteria.AddCondition("hil_productdivision", ConditionOperator.Equal, Division.Id);
        //                    query1.Criteria.AddCondition("hil_calltype", ConditionOperator.Equal, callSubType.Id);
        //                    query1.Criteria.AddCondition("hil_assignto", ConditionOperator.Equal, assignTo.Id);
        //                    EntityCollection results1 = service.RetrieveMultiple(query1);
        //                    if (results1.Entities.Count == 1)
        //                    {
        //                        foreach (Characteristic Et in results1.Entities)
        //                        {
        //                            switch (repairResource)
        //                            {
        //                                case 1:
        //                                    EntityReference AssignToAc1 = Et.hil_franchiseedirectengineer;
        //                                    Account Acc2 = (Account)service.Retrieve(Account.EntityLogicalName, AssignToAc1.Id, new ColumnSet("ownerid"));
        //                                    EntityReference AssignTo1 = Acc2.OwnerId;
        //                                    Helper.Assign(SystemUser.EntityLogicalName, msdyn_workorder.EntityLogicalName, AssignTo1.Id, serviceTicket.Id, service);
        //                                    ServiceTicket.hil_Characterstics = new EntityReference(Characteristic.EntityLogicalName, Et.Id);
        //                                    service.Update(ServiceTicket);
        //                                    break;
        //                                case 2:
        //                                    EntityReference AssignToAc = Et.hil_franchiseedirectengineer;
        //                                    Account Acc1 = (Account)service.Retrieve(Account.EntityLogicalName, AssignToAc.Id, new ColumnSet("ownerid"));
        //                                    EntityReference AssignTo = Acc1.OwnerId;
        //                                    SystemUser User = (SystemUser)service.Retrieve(SystemUser.EntityLogicalName, AssignTo.Id, new ColumnSet("positionid"));
        //                                    if (User.PositionId != null)
        //                                    {
        //                                        EntityReference Position = User.PositionId;
        //                                        if (Position.Name == "Direct Engineer")
        //                                        {
        //                                            Helper.Assign(SystemUser.EntityLogicalName, msdyn_workorder.EntityLogicalName, User.Id, serviceTicket.Id, service);
        //                                            ServiceTicket.hil_Characterstics = new EntityReference(Characteristic.EntityLogicalName, Et.Id);
        //                                            service.Update(ServiceTicket);
        //                                        }
        //                                        else if (Position.Name == "Franchisee")
        //                                        {
        //                                            HelperCallAllocationRouting.AssignWorkOrderToBranchHead(pinCode.Id, Division.Id, ServiceTicket.Id, service);
        //                                        }
        //                                    }
        //                                    else
        //                                    {
        //                                        throw new InvalidPluginExecutionException("Position is not mentioned in the User's record");
        //                                    }
        //                                    break;
        //                                case 3:
        //                                    EntityReference AssignTo1Ac = Et.hil_franchiseedirectengineer;
        //                                    Account Acc3 = (Account)service.Retrieve(Account.EntityLogicalName, AssignTo1Ac.Id, new ColumnSet("ownerid"));
        //                                    EntityReference AssignTo2 = Acc3.OwnerId;
        //                                    SystemUser User1 = (SystemUser)service.Retrieve(SystemUser.EntityLogicalName, AssignTo2.Id, new ColumnSet("positionid"));
        //                                    if (User1.PositionId != null)
        //                                    {
        //                                        EntityReference Position = User1.PositionId;
        //                                        if (Position.Name == "Franchisee")
        //                                        {
        //                                            Helper.Assign(SystemUser.EntityLogicalName, msdyn_workorder.EntityLogicalName, User1.Id, serviceTicket.Id, service);
        //                                            ServiceTicket.hil_Characterstics = new EntityReference(Characteristic.EntityLogicalName, Et.Id);
        //                                            service.Update(ServiceTicket);
        //                                        }
        //                                        else if (Position.Name == "Direct Engineer")
        //                                        {
        //                                            HelperCallAllocationRouting.AssignWorkOrderToBranchHead(pinCode.Id, Division.Id, ServiceTicket.Id, service);
        //                                        }
        //                                    }
        //                                    else
        //                                    {
        //                                        throw new InvalidPluginExecutionException("Position is not mentioned in the User's record");
        //                                    }
        //                                    break;
        //                            }

        //                        }
        //                    }
        //                    else if (results1.Entities.Count > 1)
        //                    {
        //                        bool IfFranchisee = false;
        //                        bool IfDirectEngineer = false;
        //                        Guid _DirectEng = new Guid();
        //                        Guid _Frncisee = new Guid();
        //                        foreach (Characteristic Et in results1.Entities)
        //                        {
        //                            EntityReference AssignToAcc = new EntityReference(Account.EntityLogicalName);
        //                            EntityReference AssignTo = new EntityReference(SystemUser.EntityLogicalName);
        //                            EntityReference Position = new EntityReference();
        //                            Account Acc = new Account();
        //                            SystemUser User = new SystemUser();
        //                            switch (repairResource)
        //                            {
        //                                case 1: //Both
        //                                    HelperCallAllocationRouting.AssignWorkOrderToBranchHead(pinCode.Id, Division.Id, ServiceTicket.Id, service);
        //                                    break;
        //                                case 2: //Direct Engineer
        //                                    AssignToAcc = Et.hil_franchiseedirectengineer;
        //                                    Acc = (Account)service.Retrieve(Account.EntityLogicalName, AssignToAcc.Id, new ColumnSet("ownerid"));
        //                                    AssignTo = Acc.OwnerId;
        //                                    User = (SystemUser)service.Retrieve(SystemUser.EntityLogicalName, AssignTo.Id, new ColumnSet("positionid"));
        //                                    if (User.PositionId != null)
        //                                    {
        //                                        Position = User.PositionId;
        //                                        if (Position.Name == "Direct Engineer")
        //                                        {
        //                                            IfDirectEngineer = true;
        //                                            _DirectEng = AssignTo.Id;
        //                                            CharactersticId = Et.Id;
        //                                        }
        //                                        else if (Position.Name == "Franchisee")
        //                                        {
        //                                            IfFranchisee = true;
        //                                            _Frncisee = AssignTo.Id;
        //                                        }
        //                                    }
        //                                    break;
        //                                case 3: //Franchisee
        //                                    AssignToAcc = Et.hil_franchiseedirectengineer;
        //                                    Acc = (Account)service.Retrieve(Account.EntityLogicalName, AssignToAcc.Id, new ColumnSet("ownerid"));
        //                                    AssignTo = Acc.OwnerId;
        //                                    User = (SystemUser)service.Retrieve(SystemUser.EntityLogicalName, AssignTo.Id, new ColumnSet("positionid"));
        //                                    if (User.PositionId != null)
        //                                    {
        //                                        Position = User.PositionId;
        //                                        if (Position.Name == "Direct Engineer")
        //                                        {
        //                                            IfDirectEngineer = true;
        //                                            _DirectEng = AssignTo.Id;
        //                                        }
        //                                        else if (Position.Name == "Franchisee")
        //                                        {
        //                                            IfFranchisee = true;
        //                                            _Frncisee = AssignTo.Id;
        //                                            CharactersticId = Et.Id;
        //                                        }
        //                                    }
        //                                    break;
        //                            }
        //                        }
        //                        ExecuteConditionalAssign(service, IfDirectEngineer, IfFranchisee, repairResource, _DirectEng, _Frncisee, ServiceTicket.Id, pinCode.Id, Division.Id, CharactersticId, ServiceTicket);
        //                    }
        //                    else
        //                    {
        //                        HelperCallAllocationRouting.AssignWorkOrderToBranchHead(pinCode.Id, Division.Id, ServiceTicket.Id, service);
        //                        //Helper.Assign(SystemUser.EntityLogicalName, msdyn_workorder.EntityLogicalName, Et.hil_assignto.Id, serviceTicket.Id, service);
        //                    }
        //                }
        //            }
        //        }

        //    }

        //}
        //public static void ExecuteConditionalAssign(IOrganizationService service, bool IfDirectEngineer, bool IfFranchisee, int repairResource, Guid DE,
        //    Guid Frnchisee, Guid ServiceTicket, Guid PinCode, Guid Division, Guid CharactersticId, msdyn_workorder _enServiceTicket)
        //{
        //    if (repairResource == 2)
        //    {
        //        if ((IfDirectEngineer == true) && (IfFranchisee == true))
        //        {
        //            Helper.Assign(SystemUser.EntityLogicalName, msdyn_workorder.EntityLogicalName, DE, ServiceTicket, service);
        //            if (CharactersticId != Guid.Empty)
        //                _enServiceTicket.hil_Characterstics = new EntityReference(Characteristic.EntityLogicalName, CharactersticId);
        //            service.Update(_enServiceTicket);

        //        }
        //        else if ((IfDirectEngineer == true) && (IfFranchisee == false))
        //        {
        //            HelperCallAllocationRouting.AssignWorkOrderToBranchHead(PinCode, Division, ServiceTicket, service);
        //        }
        //        else if ((IfDirectEngineer == false) && (IfFranchisee == true))
        //        {
        //            HelperCallAllocationRouting.AssignWorkOrderToBranchHead(PinCode, Division, ServiceTicket, service);
        //        }
        //    }
        //    else if (repairResource == 3)
        //    {
        //        if ((IfDirectEngineer == true) && (IfFranchisee == true))
        //        {
        //            Helper.Assign(SystemUser.EntityLogicalName, msdyn_workorder.EntityLogicalName, Frnchisee, ServiceTicket, service);
        //            if (CharactersticId != Guid.Empty)
        //                _enServiceTicket.hil_Characterstics = new EntityReference(Characteristic.EntityLogicalName, CharactersticId);
        //            service.Update(_enServiceTicket);
        //        }
        //        else if ((IfDirectEngineer == true) && (IfFranchisee == false))
        //        {
        //            HelperCallAllocationRouting.AssignWorkOrderToBranchHead(PinCode, Division, ServiceTicket, service);
        //        }
        //        else if ((IfDirectEngineer == false) && (IfFranchisee == true))
        //        {
        //            HelperCallAllocationRouting.AssignWorkOrderToBranchHead(PinCode, Division, ServiceTicket, service);
        //        }
        //    }
        //}
        #endregion
        #region Create Warranty Records
        public static void CreateWarrantyRecords(IOrganizationService service, Entity entity)
        {
            if (entity.Attributes.Contains("hil_branchheadapproval"))
            {
                OptionSetValue Approval = (OptionSetValue)entity["hil_branchheadapproval"];
                if (Approval.Value == 1)//Approved
                {
                    Entity WorkOd = service.Retrieve(msdyn_workorder.EntityLogicalName, entity.Id, new ColumnSet("msdyn_customerasset"));
                    if (WorkOd.Attributes.Contains("msdyn_customerasset"))
                    {
                        EntityReference AsstLookup = (EntityReference)WorkOd["msdyn_customerasset"];
                        Entity Asst = service.Retrieve(msdyn_customerasset.EntityLogicalName, AsstLookup.Id, new ColumnSet(true));
                        //HelperWarranty.Warranty_InitFun(Asst, service);
                        HelperWarrantyModule.Init_Warranty(Asst.Id, service);
                        msdyn_customerasset iUpdateAsset = new msdyn_customerasset();
                        iUpdateAsset.Id = Asst.Id;
                        Asst["statuscode"] = new OptionSetValue(1);
                        service.Update(Asst);
                    }
                }
            }
        }
        #endregion
        #region Check Parent Job
        public static void CheckParentJob(IOrganizationService service, Entity entity)
        {
            msdyn_workorder ServiceTicket = entity.ToEntity<msdyn_workorder>();
            if (ServiceTicket.msdyn_ParentWorkOrder != null)
            {
                Plugins.ServiceTicket.PostCreate.CheckParentJobTagged(service, entity);
            }
        }
        #endregion
        #region Calculate Charges
        public static void CalculateCharges(IOrganizationService service, Guid TicketId, msdyn_workorder preImageWo)
        {
            decimal TotalCharges = 0;
            msdyn_workorder Ticket = new msdyn_workorder();//new msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, TicketId, new ColumnSet(false));
            Ticket.Id = TicketId;
            QueryExpression Query = new QueryExpression(hil_estimate.EntityLogicalName);
            Query.ColumnSet = new ColumnSet("hil_totalcharges");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("hil_job", ConditionOperator.Equal, TicketId);
            EntityCollection Found1 = service.RetrieveMultiple(Query);
            if (false)//Found1.Entities.Count > 0
            {
                hil_estimate iEstimate = Found1.Entities[0].ToEntity<hil_estimate>();
                if (iEstimate.hil_totalCharges != null)
                {
                    Ticket.hil_actualcharges = new Money(iEstimate.hil_totalCharges.Value);
                    Ticket["hil_payblechargedecimal"] = (decimal)iEstimate.hil_totalCharges;
                    Ticket.hil_JobClosuredon = (DateTime)DateTime.Now.AddMinutes(330);
                    Ticket.msdyn_TimeClosed = (DateTime)DateTime.Now.AddMinutes(330);
                    if (!preImageWo.Contains("hil_tattime"))
                    {
                        Ticket["hil_tattime"] = (DateTime)DateTime.Now.AddMinutes(330);
                        Ticket["hil_requesttype"] = new OptionSetValue(3);
                    }
                    if (iEstimate.hil_totalCharges > 0)
                        Ticket["hil_ischargable"] = true;
                    else
                        Ticket["hil_ischargable"] = false;
                    //Ticket.hil_JobClosuredon = (DateTime)DateTime.Now.AddMinutes(330);
                    service.Update(Ticket);
                }
                else
                {
                    if (iEstimate.hil_ServiceCharges != null)
                    {
                        TotalCharges = TotalCharges + iEstimate.hil_ServiceCharges.Value;
                    }
                    if (iEstimate.hil_PartCharges != null)
                    {
                        TotalCharges = TotalCharges + iEstimate.hil_PartCharges.Value;
                    }
                    Ticket.hil_actualcharges = new Money(TotalCharges);
                    Ticket["hil_payblechargedecimal"] = (decimal)TotalCharges;
                    if (TotalCharges > 0)
                        Ticket["hil_ischargable"] = true;
                    else
                        Ticket["hil_ischargable"] = false;
                    Ticket.hil_JobClosuredon = (DateTime)DateTime.Now.AddMinutes(330);
                    Ticket.msdyn_TimeClosed = (DateTime)DateTime.Now.AddMinutes(330);
                    if (!preImageWo.Contains("hil_tattime"))
                    {
                        Ticket["hil_tattime"] = (DateTime)DateTime.Now.AddMinutes(330);
                        Ticket["hil_requesttype"] = new OptionSetValue(3);
                    }
                    service.Update(Ticket);
                }
            }
            else
            {
                QueryByAttribute Qry = new QueryByAttribute(msdyn_workorderincident.EntityLogicalName);
                Qry.ColumnSet = new ColumnSet("hil_warrantystatus", "msdyn_customerasset", "statecode");
                Qry.AddAttributeValue("msdyn_workorder", TicketId);
                EntityCollection Found = service.RetrieveMultiple(Qry);
                if (Found.Entities.Count >= 1)
                {
                    foreach (msdyn_workorderincident Inc in Found.Entities)
                    {
                        if (Inc.statecode.Equals(msdyn_workorderincidentState.Active))
                        {
                            Money TempCharges = CalculateActualChargesOutWarranty(service, Inc.Id, TicketId); ;
                            TotalCharges = TotalCharges + TempCharges.Value;
                        }
                    }
                }
                Ticket.hil_actualcharges = new Money(TotalCharges);
                Ticket["hil_payblechargedecimal"] = (decimal)TotalCharges;
                if (TotalCharges > 0)
                    Ticket["hil_ischargable"] = true;
                else
                    Ticket["hil_ischargable"] = false;
                Ticket.hil_JobClosuredon = (DateTime)DateTime.Now.AddMinutes(330);

                //================ Commented by Kuldeep Khare 31/Dec/2019 =========
                //Ticket.msdyn_TimeClosed = (DateTime)DateTime.Now.AddMinutes(330);
                //================ END ============================================
                if (!preImageWo.Contains("hil_tattime"))
                {
                    Ticket["hil_tattime"] = (DateTime)DateTime.Now.AddMinutes(330);
                    Ticket["hil_requesttype"] = new OptionSetValue(3);
                }
                service.Update(Ticket);
            }
        }
        public static Money CalculateChargesInWarranty(IOrganizationService service, Guid TemplateId, Guid IncidentId)
        {
            Money Total = new Money();
            Money PartTotal = CalculatePartChargesInWarranty(service, TemplateId, IncidentId);
            Money LaborTotal = CalculateLaborChargesInWarranty(service, TemplateId, IncidentId);
            decimal Tot = PartTotal.Value + LaborTotal.Value;
            Total = new Money(Tot);
            return Total;
        }
        public static Money CalculateActualChargesOutWarranty(IOrganizationService service, Guid IncidentId, Guid TicketId)
        {
            Money Total = new Money();
            Money ServicesTotal = CalculateServicesTotalOutWarranty(service, IncidentId, TicketId);
            Money PartsTotal = CalculatePartsTotalOutWarranty(service, IncidentId, TicketId);
            decimal Tot = ServicesTotal.Value + PartsTotal.Value;
            Total = new Money(Tot);
            return Total;
        }
        public static Money CalculateServicesTotalOutWarranty(IOrganizationService service, Guid IncidentId, Guid TicketId)
        {
            Money Total = new Money();
            decimal Mid = 0;
            QueryByAttribute Qry = new QueryByAttribute(msdyn_workorderservice.EntityLogicalName);
            Qry.ColumnSet = new ColumnSet("msdyn_service", "hil_warrantystatus");
            Qry.AddAttributeValue("msdyn_workorderincident", IncidentId);
            Qry.AddAttributeValue("msdyn_workorder", TicketId);
            Qry.AddAttributeValue("msdyn_linestatus", 690970001);
            Qry.AddAttributeValue("hil_warrantystatus", 2);
            EntityCollection Found = service.RetrieveMultiple(Qry);
            if (Found.Entities.Count >= 1)
            {
                foreach (msdyn_workorderservice Srvc in Found.Entities)
                {
                    if (Srvc.msdyn_Service != null)
                    {
                        if (Srvc.hil_WarrantyStatus.Value != 1)
                        {
                            Product ThisProduct = (Product)service.Retrieve(Product.EntityLogicalName, Srvc.msdyn_Service.Id, new ColumnSet("hil_amount"));
                            if (ThisProduct.hil_Amount != null)
                            {
                                Money Charge = ThisProduct.hil_Amount;
                                Mid = Mid + Charge.Value;
                                // Added By Kuldeep Khare 31/Dec/2019 to Calculate Effective Amount of Job Service in case of Out Warranty
                                msdyn_workorderservice Srvc1 = new msdyn_workorderservice();
                                Srvc1.Id = Srvc.Id;
                                Srvc1["hil_effectivecharge"] = Charge.Value;
                                service.Update(Srvc1);
                            }
                        }
                    }
                }
            }
            Total = new Money(Mid);
            return Total;
        }
        public static Money CalculatePartsTotalOutWarranty(IOrganizationService service, Guid IncidentId, Guid TicketId)
        {
            decimal Total = 0;
            Money TotalAmount = new Money();
            QueryByAttribute Qry = new QueryByAttribute(msdyn_workorderproduct.EntityLogicalName);
            Qry.ColumnSet = new ColumnSet("hil_replacedpart", "msdyn_product", "hil_warrantystatus", "msdyn_quantity");
            Qry.AddAttributeValue("msdyn_workorderincident", IncidentId);
            Qry.AddAttributeValue("msdyn_workorder", TicketId);
            Qry.AddAttributeValue("hil_linestatus", 910590000);
            Qry.AddAttributeValue("hil_warrantystatus", 2);
            EntityCollection Found = service.RetrieveMultiple(Qry);
            if (Found.Entities.Count >= 1)
            {
                foreach (msdyn_workorderproduct WoPdt in Found.Entities)
                {
                    if (WoPdt.msdyn_Product != null)
                    {
                        if (WoPdt.hil_WarrantyStatus.Value != 1)
                        {
                            Product ThisProduct = (Product)service.Retrieve(Product.EntityLogicalName, WoPdt.hil_replacedpart.Id, new ColumnSet("hil_amount"));
                            if (ThisProduct.hil_Amount != null && WoPdt.hil_replacedpart != null)
                            {
                                decimal CalculatedCharge = Convert.ToDecimal(ThisProduct.hil_Amount.Value * Convert.ToDecimal(WoPdt.msdyn_Quantity));
                                Money Charge = new Money(CalculatedCharge);
                                Total = Total + Charge.Value;
                                // Added By Kuldeep Khare 31/Dec/2019 to Calculate Effective Amount of Job Product in case of Out Warranty
                                msdyn_workorderproduct WoPdt1 = new msdyn_workorderproduct();
                                WoPdt1.Id = WoPdt.Id;
                                WoPdt1["hil_effectiveamount"] = CalculatedCharge;
                                service.Update(WoPdt1);
                            }
                        }
                    }
                }
            }
            TotalAmount = new Money(Total);
            return TotalAmount;
        }
        public static Money CalculatePartChargesInWarranty(IOrganizationService service, Guid TemplateId, Guid IncidentId)
        {
            decimal Total = 0;
            Money total = new Money();
            QueryByAttribute Query = new QueryByAttribute(hil_part.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(true);
            Query.AddAttributeValue("hil_warrantytemplateid", TemplateId);
            Query.AddAttributeValue("hil_includedinwarranty", 2);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count >= 1)
            {
                foreach (hil_part Part in Found.Entities)
                {
                    if (Part.hil_PartCode != null)
                    {
                        bool CheckIfEstimated = CheckIfPartUsed(service, Part.hil_PartCode.Id, IncidentId);
                        if (CheckIfEstimated == true)
                        {
                            decimal GetQuantity = GetQuantityFromJobProduct(service, Part.hil_PartCode.Id, IncidentId);
                            Product Pdt = (Product)service.Retrieve(Product.EntityLogicalName, Part.hil_PartCode.Id, new ColumnSet("hil_amount"));
                            Total = Total + (Pdt.hil_Amount.Value * GetQuantity);
                        }
                    }
                }
            }
            total = new Money(Total);
            return total;
        }
        public static decimal GetQuantityFromJobProduct(IOrganizationService service, Guid Part, Guid Incident)
        {
            decimal Quantity = 1;
            QueryByAttribute query = new QueryByAttribute(msdyn_workorderproduct.EntityLogicalName);
            query.ColumnSet = new ColumnSet("msdyn_quantity");
            query.AddAttributeValue("msdyn_workorderincident", Incident);
            query.AddAttributeValue("msdyn_product", Part);
            query.AddAttributeValue("msdyn_linestatus", 690970001);
            EntityCollection Found = service.RetrieveMultiple(query);
            if (Found.Entities.Count >= 1)
            {
                foreach (msdyn_workorderproduct WoPdt in Found.Entities)
                {
                    Quantity = Convert.ToDecimal(WoPdt.msdyn_Quantity);
                }
            }
            return Quantity;
        }
        public static Money CalculateLaborChargesInWarranty(IOrganizationService service, Guid TemplateId, Guid IncidentId)
        {
            decimal Total = 0;
            Money total = new Money();
            QueryByAttribute Query = new QueryByAttribute(hil_labor.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(true);
            Query.AddAttributeValue("hil_warrantytemplateid", TemplateId);
            Query.AddAttributeValue("hil_includedinwarranty", 2);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count >= 1)
            {
                foreach (hil_labor Labor in Found.Entities)
                {
                    if (Labor.hil_Labor != null)
                    {

                        bool CheckIfEstimated = CheckIfLaborUsed(service, Labor.hil_Labor.Id, IncidentId);
                        if (CheckIfEstimated == true)
                        {
                            Product Pdt = (Product)service.Retrieve(Product.EntityLogicalName, Labor.hil_Labor.Id, new ColumnSet("hil_amount"));
                            Total = Total + Pdt.hil_Amount.Value;
                        }
                    }
                }
            }
            total = new Money(Total);
            return total;
        }
        public static bool CheckIfPartUsed(IOrganizationService service, Guid Part, Guid Incident)
        {
            bool MarkIfUsed = false;
            QueryByAttribute query = new QueryByAttribute(msdyn_workorderproduct.EntityLogicalName);
            query.ColumnSet = new ColumnSet(false);
            query.AddAttributeValue("msdyn_workorderincident", Incident);
            query.AddAttributeValue("msdyn_product", Part);
            query.AddAttributeValue("msdyn_linestatus", 690970001);
            EntityCollection Found = service.RetrieveMultiple(query);
            if (Found.Entities.Count >= 1)
            {
                foreach (msdyn_workorderproduct Pdt in Found.Entities)
                {
                    if (Pdt.hil_replacedpart != null)
                    {
                        MarkIfUsed = true;
                    }
                }
            }
            return MarkIfUsed;
        }
        public static bool CheckIfLaborUsed(IOrganizationService service, Guid Labor, Guid Incident)
        {
            bool MarkIfUsed = false;
            msdyn_workorderservice Serv = new msdyn_workorderservice();
            QueryByAttribute query = new QueryByAttribute(msdyn_workorderservice.EntityLogicalName);
            query.ColumnSet = new ColumnSet(false);
            query.AddAttributeValue("msdyn_workorderincident", Incident);
            query.AddAttributeValue("msdyn_service", Labor);
            query.AddAttributeValue("msdyn_linestatus", 690970001);
            EntityCollection Found = service.RetrieveMultiple(query);
            if (Found.Entities.Count >= 1)
            {
                MarkIfUsed = true;
            }
            return MarkIfUsed;
        }
        #endregion
        #region Set Warranty Status
        public static void SetWarrantyStatus(IOrganizationService service, msdyn_workorder Wo)
        {
            tracingService.Trace("6");
            msdyn_workorder Wo1 = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, Wo.Id, new ColumnSet("hil_productsubcategory"));
            msdyn_workorder enUpdateWo = new msdyn_workorder();
            enUpdateWo.Id = Wo1.Id;
            if (Wo.hil_PurchaseDate != null && Wo1.hil_ProductSubcategory != null)
            {
                tracingService.Trace("7");
                DateTime date1 = (DateTime)Wo.hil_PurchaseDate;
                DateTime date2 = DateTime.Now;
                int Diff = ((date2.Year - date1.Year) * 12) + (date2.Month - date1.Month) * 30 + date2.Day - date1.Day;
                QueryByAttribute Qry = new QueryByAttribute(hil_warrantytemplate.EntityLogicalName);
                Qry.ColumnSet = new ColumnSet(true);
                Qry.AddAttributeValue("hil_product", Wo1.hil_ProductSubcategory.Id);
                EntityCollection Found = service.RetrieveMultiple(Qry);
                if (Found.Entities.Count >= 1)
                {
                    foreach (hil_warrantytemplate Wtmp in Found.Entities)
                    {
                        int Period = (int)(Wtmp.hil_WarrantyPeriod * 30);
                        if (Period >= Diff)
                        {
                            enUpdateWo.hil_WarrantyStatus = new OptionSetValue(1);
                        }
                        else
                        {
                            enUpdateWo.hil_WarrantyStatus = new OptionSetValue(2);
                        }
                    }
                }
                //else if (Found.Entities.Count > 1)
                //{
                //    Wo1.hil_WarrantyStatus = new OptionSetValue(3);
                //}
                else
                {
                    enUpdateWo.hil_WarrantyStatus = new OptionSetValue(3);
                }
                service.Update(enUpdateWo);
            }
        }
        #endregion
        #region Estimate Charges 
        public static void EstimateCharges(IOrganizationService service, Guid TicketId)
        {
            decimal TotalCharges = 0;
            msdyn_workorder Ticket = new msdyn_workorder();//(msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, TicketId, new ColumnSet(false));
            Ticket.Id = TicketId;
            QueryByAttribute Qry = new QueryByAttribute(msdyn_workorderincident.EntityLogicalName);
            Qry.ColumnSet = new ColumnSet("hil_warrantystatus", "msdyn_customerasset", "statecode");
            Qry.AddAttributeValue("msdyn_workorder", TicketId);
            EntityCollection Found = service.RetrieveMultiple(Qry);
            if (Found.Entities.Count >= 1)
            {
                foreach (msdyn_workorderincident Inc in Found.Entities)
                {
                    if (Inc.statecode.Equals(msdyn_workorderincidentState.Active))
                    {
                        Money TempCharges = CalculateActualChargesEstimateOutWarranty(service, Inc.Id, TicketId); ;
                        TotalCharges = TotalCharges + TempCharges.Value;
                    }
                }
            }
            Ticket.hil_EstimateChargesTotal = new Money(TotalCharges);
            Ticket["hil_estimatedchargedecimal"] = (decimal)TotalCharges;
            service.Update(Ticket);
        }
        public static Money CalculateChargesInWarrantyEstimate(IOrganizationService service, Guid TemplateId, Guid Incident)
        {
            Money Total = new Money();
            Money PartTotal = CalculatePartChargesInWarrantyEstimate(service, TemplateId, Incident);
            Money LaborTotal = CalculateLaborChargesInWarrantyEstimate(service, TemplateId, Incident);
            decimal Tot = PartTotal.Value + LaborTotal.Value;
            Total = new Money(Tot);
            return Total;
        }
        public static Money CalculateActualChargesEstimateOutWarranty(IOrganizationService service, Guid IncidentId, Guid TicketId)
        {
            Money Total = new Money();
            Money ServicesTotal = CalculateServicesTotalEstimateOutWarranty(service, IncidentId, TicketId);
            Money PartsTotal = CalculatePartsTotalEstimateOutWarranty(service, IncidentId, TicketId);
            decimal Tot = ServicesTotal.Value + PartsTotal.Value;
            Total = new Money(Tot);
            return Total;
        }
        public static Money CalculateServicesTotalEstimateOutWarranty(IOrganizationService service, Guid IncidentId, Guid TicketId)
        {
            Money Total = new Money();
            decimal Mid = 0;
            QueryByAttribute Qry = new QueryByAttribute(msdyn_workorderservice.EntityLogicalName);
            Qry.ColumnSet = new ColumnSet(true);
            Qry.AddAttributeValue("msdyn_workorderincident", IncidentId);
            Qry.AddAttributeValue("msdyn_workorder", TicketId);
            //Qry.AddAttributeValue("msdyn_linestatus", 690970000);
            //Qry.AddAttributeValue("hil_warrantystatus", 2);
            EntityCollection Found = service.RetrieveMultiple(Qry);
            if (Found.Entities.Count >= 1)
            {
                foreach (msdyn_workorderservice Srvc in Found.Entities)
                {
                    if (Srvc.hil_WarrantyStatus.Value != 1 && Srvc.msdyn_LineStatus != null)
                    {
                        if (Srvc.msdyn_Service != null)
                        {
                            Product ThisProduct = (Product)service.Retrieve(Product.EntityLogicalName, Srvc.msdyn_Service.Id, new ColumnSet("hil_amount"));
                            if (ThisProduct.hil_Amount != null)
                            {
                                Money Charge = ThisProduct.hil_Amount;
                                Mid = Mid + Charge.Value;
                            }
                        }
                    }
                }
            }
            Total = new Money(Mid);
            return Total;
        }
        public static Money CalculatePartsTotalEstimateOutWarranty(IOrganizationService service, Guid IncidentId, Guid TicketId)
        {
            decimal Total = 0;
            Money TotalAmount = new Money();
            QueryByAttribute Qry = new QueryByAttribute(msdyn_workorderproduct.EntityLogicalName);
            Qry.ColumnSet = new ColumnSet(true);
            Qry.AddAttributeValue("msdyn_workorderincident", IncidentId);
            Qry.AddAttributeValue("msdyn_workorder", TicketId);
            //Qry.AddAttributeValue("msdyn_linestatus", 690970000);
            //Qry.AddAttributeValue("hil_warrantystatus", 2);
            EntityCollection Found = service.RetrieveMultiple(Qry);
            if (Found.Entities.Count >= 1)
            {
                foreach (msdyn_workorderproduct WoPdt in Found.Entities)
                {
                    if (WoPdt.msdyn_Product != null && WoPdt.hil_replacedpart != null)
                    {
                        if (WoPdt.hil_WarrantyStatus.Value != 1)
                        {
                            Product ThisProduct = (Product)service.Retrieve(Product.EntityLogicalName, WoPdt.msdyn_Product.Id, new ColumnSet("hil_amount"));
                            if (ThisProduct.hil_Amount != null)
                            {
                                Money Charge = new Money(ThisProduct.hil_Amount.Value * Convert.ToDecimal(WoPdt.msdyn_Quantity));
                                Total = Total + Charge.Value;
                            }
                        }
                    }
                }
            }
            TotalAmount = new Money(Total);
            return TotalAmount;
        }
        public static Money CalculatePartChargesInWarrantyEstimate(IOrganizationService service, Guid TemplateId, Guid Incident)
        {
            decimal Total = 0;
            Money total = new Money();
            QueryByAttribute Query = new QueryByAttribute(hil_part.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(true);
            Query.AddAttributeValue("hil_warrantytemplateid", TemplateId);
            Query.AddAttributeValue("hil_includedinwarranty", 2);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count >= 1)
            {
                foreach (hil_part Part in Found.Entities)
                {
                    if (Part.hil_PartCode != null)
                    {
                        bool IfEstimated = CheckIfEstimatedPart(service, Part.hil_PartCode.Id, Incident);
                        if (IfEstimated == true)
                        {
                            decimal GetQuantity = GetQuantityFromJobProduct(service, Part.hil_PartCode.Id, Incident);
                            Product Pdt = (Product)service.Retrieve(Product.EntityLogicalName, Part.hil_PartCode.Id, new ColumnSet("hil_amount"));
                            Total = Total + Pdt.hil_Amount.Value * GetQuantity;
                        }
                    }
                }
            }
            total = new Money(Total);
            return total;
        }
        public static Money CalculateLaborChargesInWarrantyEstimate(IOrganizationService service, Guid TemplateId, Guid Incident)
        {
            decimal Total = 0;
            Money total = new Money();
            QueryByAttribute Query = new QueryByAttribute(hil_labor.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(true);
            Query.AddAttributeValue("hil_warrantytemplateid", TemplateId);
            Query.AddAttributeValue("hil_includedinwarranty", 2);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count >= 1)
            {
                foreach (hil_labor Labor in Found.Entities)
                {
                    if (Labor.hil_Labor != null)
                    {
                        bool IfEstimated = CheckIfEstimatedLabor(service, Labor.hil_Labor.Id, Incident);
                        if (IfEstimated == true)
                        {
                            Product Pdt = (Product)service.Retrieve(Product.EntityLogicalName, Labor.hil_Labor.Id, new ColumnSet("hil_amount"));
                            Total = Total + Pdt.hil_Amount.Value;
                        }
                    }
                }
            }
            total = new Money(Total);
            return total;
        }
        public static bool CheckIfEstimatedLabor(IOrganizationService service, Guid Labor, Guid Incident)
        {
            bool IfEstimatedLabor = false;
            QueryByAttribute Query = new QueryByAttribute(msdyn_workorderservice.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(false);
            Query.AddAttributeValue("msdyn_workorderincident", Incident);
            Query.AddAttributeValue("msdyn_service", Labor);
            Query.AddAttributeValue("msdyn_linestatus", 690970000);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count >= 1)
            {
                IfEstimatedLabor = true;
            }
            return IfEstimatedLabor;
        }
        public static bool CheckIfEstimatedPart(IOrganizationService service, Guid Part, Guid Incident)
        {
            bool IfEstimatedPart = false;
            QueryByAttribute Query = new QueryByAttribute(msdyn_workorderproduct.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(false);
            Query.AddAttributeValue("msdyn_workorderincident", Incident);
            Query.AddAttributeValue("msdyn_product", Part);
            Query.AddAttributeValue("msdyn_linestatus", 690970000);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count >= 1)
            {
                foreach (msdyn_workorderproduct Pdt in Found.Entities)
                {
                    if (Pdt.hil_replacedpart != null)
                    {
                        IfEstimatedPart = true;
                    }
                }
            }
            return IfEstimatedPart;
        }
        #endregion
        #region Create Warranty for Replaced Parts
        public static void CreateWarrantyForReplacedPart(IOrganizationService service, Guid JobId)
        {
            QueryExpression Query = new QueryExpression(msdyn_workorderproduct.EntityLogicalName);
            Query.ColumnSet = new ColumnSet("msdyn_product", "hil_replacedpart", "hil_warrantystatus", "msdyn_customerasset");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("msdyn_workorder", ConditionOperator.Equal, JobId);
            Query.Criteria.AddCondition("msdyn_linestatus", ConditionOperator.Equal, 690970001);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                foreach (msdyn_workorderproduct _jobPdt in Found.Entities)
                {
                    if ((_jobPdt.hil_WarrantyStatus != null) && (_jobPdt.hil_replacedpart != null))
                    {
                        if (_jobPdt.hil_WarrantyStatus.Value == 1)
                        {
                            UpdateUnitWarrantyForThisPart(service, _jobPdt);
                        }
                        else if ((_jobPdt.hil_WarrantyStatus.Value == 2) || (_jobPdt.hil_WarrantyStatus.Value == 3))
                        {
                            CreateUnitWarrantyforReplacedPart(service, _jobPdt);
                        }
                    }
                }
            }
        }
        public static void CreateUnitWarrantyforReplacedPart(IOrganizationService service, msdyn_workorderproduct _jobPdt)
        {
            hil_unitwarranty _enWty = new hil_unitwarranty();
            msdyn_customerasset Asset = (msdyn_customerasset)service.Retrieve(msdyn_customerasset.EntityLogicalName, _jobPdt.msdyn_CustomerAsset.Id, new ColumnSet("hil_productcategory", "hil_productsubcategory"));
            EntityReference _wtyTmp = GetWarrantyTemplate(service, Asset);
            _enWty.hil_warrantystartdate = DateTime.Now;
            _enWty.hil_warrantyenddate = DateTime.Now.AddDays(30);
            _enWty.hil_Part = _jobPdt.hil_replacedpart;
            _enWty.hil_productitem = Asset.hil_ProductSubcategory;
            _enWty.hil_productmodel = Asset.hil_ProductCategory;
            _enWty.hil_CustomerAsset = new EntityReference(msdyn_customerasset.EntityLogicalName, Asset.Id);
            _enWty.hil_ProductType = new OptionSetValue(2);
            if (_wtyTmp != null)
            {
                _enWty.hil_WarrantyTemplate = _wtyTmp;
            }
            service.Create(_enWty);
        }
        public static EntityReference GetWarrantyTemplate(IOrganizationService service, msdyn_customerasset Asset)
        {
            EntityReference _enTemp = new EntityReference();
            QueryByAttribute Query = new QueryByAttribute(hil_warrantytemplate.EntityLogicalName);
            Query.AddAttributeValue("hil_product", Asset.hil_ProductSubcategory.Id);
            Query.AddAttributeValue("hil_type", 1);
            ColumnSet Columns = new ColumnSet(false);
            Query.ColumnSet = Columns;
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count >= 1)
            {
                hil_warrantytemplate Tmp = (hil_warrantytemplate)Found.Entities[0];
                _enTemp = new EntityReference(hil_warrantytemplate.EntityLogicalName, Tmp.Id);
            }
            return _enTemp;
        }
        public static void UpdateUnitWarrantyForThisPart(IOrganizationService service, msdyn_workorderproduct _jobPdt)
        {
            if (_jobPdt.msdyn_CustomerAsset != null && _jobPdt.msdyn_Product != null)
            {
                QueryByAttribute Query = new QueryByAttribute(hil_unitwarranty.EntityLogicalName);
                Query.ColumnSet = new ColumnSet(false);
                Query.AddAttributeValue("hil_customerasset", _jobPdt.msdyn_CustomerAsset.Id);
                Query.AddAttributeValue("hil_part", _jobPdt.msdyn_Product.Id);
                EntityCollection Found = service.RetrieveMultiple(Query);
                if (Found.Entities.Count > 0)
                {
                    hil_unitwarranty _unWt = (hil_unitwarranty)Found.Entities[0];
                    _unWt.hil_Part = _jobPdt.hil_replacedpart;
                    service.Update(_unWt);
                }
            }
        }
        #endregion
        #region Installation Date on Asset
        public static void InstallationDateOnAsset(IOrganizationService service, Guid JobId)
        {
            msdyn_workorder _enJob = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, JobId, new ColumnSet("hil_callsubtype"));
            if (_enJob.hil_CallSubType != null)
            {
                if (_enJob.hil_CallSubType.Name == "Installation")
                {
                    QueryByAttribute Query = new QueryByAttribute(msdyn_workorderproduct.EntityLogicalName);
                    Query.ColumnSet = new ColumnSet("msdyn_customerasset");
                    Query.AddAttributeValue("msdyn_workorder", _enJob.Id);
                    EntityCollection Found = service.RetrieveMultiple(Query);
                    if (Found.Entities.Count > 0)
                    {
                        foreach (msdyn_workorderproduct _jPdt in Found.Entities)
                        {
                            if (_jPdt.msdyn_CustomerAsset != null)
                            {
                                msdyn_customerasset Asset = new msdyn_customerasset();
                                Asset.Id = _jPdt.msdyn_CustomerAsset.Id;//(msdyn_customerasset)service.Retrieve(msdyn_customerasset.EntityLogicalName, _jPdt.msdyn_CustomerAsset.Id, new ColumnSet(false));
                                Asset["hil_purchasedate"] = DateTime.Now;
                                service.Update(Asset);
                            }
                        }
                    }
                }
            }
        }
        /// <summary>
        /// New Method developed by Kuldeep Khare to update first Installation Information on Asset
        /// </summary>
        /// <param name="service">IOrganisation Service Proxy</param>
        /// <param name="JobId">Work Order GUID</param>
        public static void UpdateFirstInstallationInformationOnAsset(IOrganizationService service, Guid JobId)
        {
            msdyn_workorder _enJob = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, JobId, new ColumnSet("hil_callsubtype", "msdyn_substatus", "msdyn_name"));
            if (_enJob.hil_CallSubType != null)
            {
                if (_enJob.hil_CallSubType.Name == "Installation")
                {
                    QueryByAttribute Query = new QueryByAttribute(msdyn_workorderincident.EntityLogicalName);
                    Query.ColumnSet = new ColumnSet("msdyn_customerasset");
                    Query.AddAttributeValue("msdyn_workorder", _enJob.Id);
                    EntityCollection Found = service.RetrieveMultiple(Query);
                    if (Found.Entities.Count > 0)
                    {
                        foreach (msdyn_workorderincident _jPdt in Found.Entities)
                        {
                            if (_jPdt.msdyn_CustomerAsset != null)
                            {
                                #region Check whether current job is first Installation job of the Asset.
                                QueryExpression qeObj = new QueryExpression(msdyn_workorder.EntityLogicalName);
                                qeObj.ColumnSet = new ColumnSet("msdyn_workorderid", "hil_jobclosuredon");
                                qeObj.Criteria = new FilterExpression(LogicalOperator.And);
                                qeObj.Criteria.AddCondition(new ConditionExpression("msdyn_workorderid", ConditionOperator.NotEqual, JobId));
                                qeObj.Criteria.AddCondition(new ConditionExpression("hil_callsubtype", ConditionOperator.Equal, _enJob.hil_CallSubType.Id));
                                //Change the GUID @Go-Live
                                qeObj.Criteria.AddCondition(new ConditionExpression("msdyn_substatus", ConditionOperator.Equal, new Guid("1727FA6C-FA0F-E911-A94E-000D3AF060A1"))); //{Substatus:"Closed"}-
                                LinkEntity le = qeObj.AddLink("msdyn_workorderincident", "msdyn_workorderid", "msdyn_workorder", JoinOperator.Inner);
                                le.LinkCriteria = new FilterExpression(LogicalOperator.And);
                                le.LinkCriteria.AddCondition("msdyn_customerasset", ConditionOperator.Equal, _jPdt.msdyn_CustomerAsset.Id);
                                qeObj.TopCount = 1;
                                qeObj.AddOrder("hil_jobclosuredon", OrderType.Ascending);
                                EntityCollection ecObj = service.RetrieveMultiple(qeObj);
                                if (ecObj.Entities.Count <= 0)
                                {
                                    #region Update First Installation Information on Asseet
                                    msdyn_customerasset Asset = new msdyn_customerasset();
                                    Asset.Id = _jPdt.msdyn_CustomerAsset.Id;
                                    //(msdyn_customerasset)service.Retrieve(msdyn_customerasset.EntityLogicalName, _jPdt.msdyn_CustomerAsset.Id, new ColumnSet(false));
                                    Asset["hil_firstinstallationdate"] = DateTime.Now;
                                    Asset["hil_firstinstallationjob"] = _enJob.GetAttributeValue<string>("msdyn_name");
                                    service.Update(Asset);
                                    #endregion
                                }
                                #endregion
                            }
                        }
                    }
                }
            }
        }
        #endregion
        #region Check If Parts or Service Estimated
        public static bool CheckIfPartsorServiceEstimated(IOrganizationService service, Guid WoId)
        {
            bool IfPresent = false;
            QueryExpression Query = new QueryExpression(msdyn_workorderproduct.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("msdyn_workorder", ConditionOperator.Equal, WoId);
            Query.Criteria.AddCondition("msdyn_linestatus", ConditionOperator.Equal, 690970000);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count == 0)
            {
                QueryExpression Query1 = new QueryExpression(msdyn_workorderservice.EntityLogicalName);
                Query1.ColumnSet = new ColumnSet(false);
                Query1.Criteria = new FilterExpression(LogicalOperator.And);
                Query1.Criteria.AddCondition("msdyn_workorder", ConditionOperator.Equal, WoId);
                Query1.Criteria.AddCondition("msdyn_linestatus", ConditionOperator.Equal, 690970000);
                EntityCollection Found1 = service.RetrieveMultiple(Query1);
                if (Found1.Entities.Count == 0)
                {
                    IfPresent = false;
                }
                else
                {
                    IfPresent = true;
                }
            }
            else
            {
                IfPresent = true;
            }
            return IfPresent;
        }
        public static bool CheckIfReplacedPartNullOrNotAvailable(IOrganizationService service, Guid WoId)
        {
            bool IfPresent = false;
            msdyn_workorderproduct enUpdateWOProduct = new msdyn_workorderproduct();
            QueryExpression Query = new QueryExpression(msdyn_workorderproduct.EntityLogicalName);
            Query.ColumnSet = new ColumnSet("hil_replacedpart", "hil_availabilitystatus", "hil_linestatus", "hil_markused");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("msdyn_workorder", ConditionOperator.Equal, WoId);
            Query.Criteria.AddCondition("hil_replacedpart", ConditionOperator.NotNull);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                foreach (msdyn_workorderproduct _wPdt in Found.Entities)
                {
                    if (_wPdt.hil_AvailabilityStatus.Value == 2 || _wPdt.hil_LineStatus.Value == 1)//|| _wPdt.msdyn_LineStatus.Value == 690970001)
                    {
                        IfPresent = true;
                        break;
                    }
                    else
                    {
                        enUpdateWOProduct.Id = _wPdt.Id;
                        if (!_wPdt.Attributes.Contains("hil_markused"))
                        {
                            enUpdateWOProduct["hil_markused"] = true;
                            service.Update(enUpdateWOProduct);
                        }
                        else
                        {
                            bool opMarkUsed = _wPdt.GetAttributeValue<bool>("hil_markused");
                            if (opMarkUsed != true)
                            {
                                enUpdateWOProduct["hil_markused"] = true;
                                service.Update(enUpdateWOProduct);
                            }
                        }
                    }
                    //if(_wPdt.hil_AvailabilityStatus.Value == 2 || _wPdt.hil_LineStatus.Value == 1)
                    //{
                    //    IfPresent = true;
                    //    break;
                    //}
                }
            }
            return IfPresent;
        }
        #endregion
        #region Estimate Header
        public static void CreateEstimateHeader(IOrganizationService service, Guid entityId)
        {
            Guid _estHeaderId = GetEstimateHeader(service, entityId);
            if (_estHeaderId != Guid.Empty)
            {
                decimal PartAmount = 0;
                decimal ServiceAmount = 0;
                QueryExpression Query = new QueryExpression(msdyn_workorderproduct.EntityLogicalName);
                Query.ColumnSet = new ColumnSet("hil_partamount", "hil_replacedpart", "hil_warrantystatus", "statecode");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition("msdyn_workorder", ConditionOperator.Equal, entityId);
                Query.Criteria.AddCondition("hil_warrantystatus", ConditionOperator.NotEqual, 1);
                Query.Criteria.AddCondition("hil_replacedpart", ConditionOperator.NotNull);
                //Query.AddAttributeValue("hil_warrantystatus", 2);
                //Query.AddAttributeValue("hil_warrantystatus", 690970000);
                EntityCollection Found = service.RetrieveMultiple(Query);
                if (Found.Entities.Count > 0)
                {
                    foreach (msdyn_workorderproduct Pdt1 in Found.Entities)
                    {
                        if (Pdt1.hil_WarrantyStatus.Value != 1 && Pdt1.statecode.Equals(msdyn_workorderproductState.Active))
                        {
                            if (Pdt1.hil_replacedpart != null)
                            {
                                msdyn_workorderproduct Pdt = new msdyn_workorderproduct();
                                Pdt.Id = Pdt1.Id;//(msdyn_workorderproduct)service.Retrieve(msdyn_workorderproduct.EntityLogicalName, Pdt1.Id, new ColumnSet(false));
                                Pdt.hil_RegardingEstimate = new EntityReference(hil_estimate.EntityLogicalName, _estHeaderId);
                                Pdt.hil_DiscountedAmount = 0;
                                Pdt.hil_FinalAmount = Pdt1.hil_PartAmount;
                                service.Update(Pdt);
                                PartAmount = PartAmount + Convert.ToDecimal(Pdt1.hil_PartAmount);
                            }
                        }
                    }
                }
                msdyn_workorderservice enUpdatePdt = new msdyn_workorderservice();
                QueryExpression Query1 = new QueryExpression(msdyn_workorderservice.EntityLogicalName);
                Query1.ColumnSet = new ColumnSet("hil_charge", "hil_warrantystatus", "statecode");
                Query1.Criteria = new FilterExpression(LogicalOperator.And);
                Query1.Criteria.AddCondition("msdyn_workorder", ConditionOperator.Equal, entityId);
                Query1.Criteria.AddCondition("hil_warrantystatus", ConditionOperator.NotEqual, 1);
                //Query1.Criteria.AddCondition("msdyn_linestatus", ConditionOperator.Equal, 690970000);
                Query1.Criteria.AddCondition("msdyn_linestatus", ConditionOperator.NotNull);//added on 24 May
                EntityCollection Found1 = service.RetrieveMultiple(Query1);
                if (Found1.Entities.Count > 0)
                {
                    foreach (msdyn_workorderservice Pdt in Found1.Entities)
                    {
                        if (Pdt.hil_WarrantyStatus.Value != 1 && Pdt.statecode.Equals(msdyn_workorderserviceState.Active))
                        {
                            enUpdatePdt.Id = Pdt.Id;
                            enUpdatePdt.hil_RegardingEstimate = new EntityReference(hil_estimate.EntityLogicalName, _estHeaderId);
                            enUpdatePdt.hil_DiscountedAmount = 0;
                            enUpdatePdt.hil_FinalAmount = Pdt.hil_Charge;
                            service.Update(enUpdatePdt);
                            ServiceAmount = ServiceAmount + Convert.ToDecimal(Pdt.hil_Charge);
                        }
                    }
                }
                hil_estimate _estHeader = new hil_estimate();
                _estHeader.Id = _estHeaderId;//(hil_estimate)service.Retrieve(hil_estimate.EntityLogicalName, _estHeaderId, new ColumnSet(false));
                if (Found.Entities.Count == 0 && Found1.Entities.Count == 0)
                {
                    //service.Delete(hil_estimate.EntityLogicalName, _estHeaderId);
                }
                else
                {
                    _estHeader.hil_PartCharges = PartAmount;
                    _estHeader.hil_ServiceCharges = ServiceAmount;
                    _estHeader["hil_partdiscount"] = (decimal)0;
                    _estHeader["hil_servicediscount"] = (decimal)0;
                    //_estHeader["hil_branchheadapprovalstatus"] = null;
                    // _estHeader["hil_customerapprovalstatus"] = null;
                    _estHeader.statuscode = new OptionSetValue(1);
                    service.Update(_estHeader);
                }
            }
        }
        public static Guid GetEstimateHeader(IOrganizationService service, Guid entityId)
        {
            Guid HeaderId = new Guid();
            QueryByAttribute Query = new QueryByAttribute(hil_estimate.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(false);
            Query.AddAttributeValue("hil_job", entityId);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count == 0)
            {
                msdyn_workorder eJob = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, entityId, new ColumnSet("hil_productcategory", "hil_assignmentmatrix"));
                hil_estimate Header = new hil_estimate();
                Header.hil_Job = new EntityReference(msdyn_workorder.EntityLogicalName, entityId);
                if (eJob.hil_Productcategory != null)
                    Header["hil_division"] = eJob.hil_Productcategory;
                HeaderId = service.Create(Header);
                if (eJob.hil_AssignmentMatrix != null)
                {
                    hil_assignmentmatrix iMat = (hil_assignmentmatrix)service.Retrieve(hil_assignmentmatrix.EntityLogicalName, eJob.hil_AssignmentMatrix.Id, new ColumnSet("ownerid"));
                    Helper.Assign(iMat.OwnerId.LogicalName, hil_estimate.EntityLogicalName, iMat.OwnerId.Id, HeaderId, service);
                }

            }
            else
            {
                HeaderId = Found.Entities[0].Id;
                RemoveExistingPartsandServices(service, HeaderId);
            }
            return HeaderId;
        }
        public static void RemoveExistingPartsandServices(IOrganizationService service, Guid HeaderId)
        {
            QueryByAttribute Query = new QueryByAttribute(msdyn_workorderproduct.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(false);
            Query.AddAttributeValue("hil_regardingestimate", HeaderId);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                foreach (msdyn_workorderproduct _enJpD in Found.Entities)
                {
                    _enJpD.hil_RegardingEstimate = null;
                    _enJpD.hil_DiscountedAmount = null;
                    _enJpD.hil_FinalAmount = null;
                    service.Update(_enJpD);
                }
            }
            QueryByAttribute Query1 = new QueryByAttribute(msdyn_workorderservice.EntityLogicalName);
            Query1.ColumnSet = new ColumnSet(false);
            Query1.AddAttributeValue("hil_regardingestimate", HeaderId);
            EntityCollection Found1 = service.RetrieveMultiple(Query1);
            if (Found1.Entities.Count > 0)
            {
                foreach (msdyn_workorderservice _enJsV in Found1.Entities)
                {
                    _enJsV.hil_RegardingEstimate = null;
                    _enJsV.hil_DiscountedAmount = null;
                    _enJsV.hil_FinalAmount = null;
                    service.Update(_enJsV);
                }
            }
        }
        #endregion
        #region Create Phone Call
        public static void CreatePhoneCall(IOrganizationService service, Guid _jId)
        {
            msdyn_workorder _eJob = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, _jId, new ColumnSet(""));
            PhoneCall _enPhone = new PhoneCall();
            _enPhone.hil_Consumer = _eJob.hil_CustomerRef;
            _enPhone.hil_CallingNumber = _eJob.hil_CallingNumber;
            _enPhone.hil_AlternateNumber1 = _eJob.hil_Alternate;
            _enPhone.PhoneNumber = _eJob.hil_mobilenumber;
            _enPhone.hil_Disposition = new OptionSetValue(1);
            _enPhone.DirectionCode = true;
            _enPhone.RegardingObjectId = new EntityReference(msdyn_workorder.EntityLogicalName, _jId);
            service.Create(_enPhone);
        }
        #endregion
        #region Check If Direct Engineer Present
        public static bool CheckIfDirectEngineerPresent(IOrganizationService service, Guid iEngineer)
        {
            bool IfPresent = false;
            QueryExpression Query = new QueryExpression(msdyn_timeoffrequest.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition(new ConditionExpression("ownerid", ConditionOperator.Equal, iEngineer));
            Query.Criteria.AddCondition(new ConditionExpression("msdyn_starttime", ConditionOperator.OnOrBefore, DateTime.Now));
            Query.Criteria.AddCondition(new ConditionExpression("msdyn_endtime", ConditionOperator.OnOrAfter, DateTime.Now));
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                IfPresent = true;
            }
            return IfPresent;
        }
        #endregion
        #region KKG De-Code
        public static void CheckIfValidKKG(IOrganizationService service, Entity entity)
        {
            string Decrypted = string.Empty;
            string Match = string.Empty;
            msdyn_workorder iWo = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, entity.Id, new ColumnSet("hil_kkgotp"));
            msdyn_workorder PostEnt = entity.ToEntity<msdyn_workorder>();
            if (PostEnt.hil_kkgcode != null && iWo.hil_KKGOTP != null)
            {
                Match = iWo.hil_KKGOTP;
                Decrypted = DecryptOTP(PostEnt.hil_kkgcode);
                if (Match != Decrypted)
                {
                    throw new InvalidPluginExecutionException("XXXXXXXXXXXXXXXXXXXXXXXXXXX - KKG OTP DOESN'T MATCH - XXXXXXXXXXXXXXXXXXXXXXXXXXX");
                }
            }
        }
        public static string DecryptOTP(string OTP_Clean)
        {
            string strAlpha = string.Empty;
            Char Index = new char();
            for (int i = 0; i < OTP_Clean.Length; i++)
            {
                Index = OTP_Clean[i];
                switch (Index)
                {
                    case '0':
                        strAlpha = strAlpha + "S";
                        break;
                    case '1':
                        strAlpha = strAlpha + "A";
                        break;
                    case '2':
                        strAlpha = strAlpha + "M";
                        break;
                    case '3':
                        strAlpha = strAlpha + "Q";
                        break;
                    case '4':
                        strAlpha = strAlpha + "Z";
                        break;
                    case '5':
                        strAlpha = strAlpha + "T";
                        break;
                    case '6':
                        strAlpha = strAlpha + "P";
                        break;
                    case '7':
                        strAlpha = strAlpha + "K";
                        break;
                    case '8':
                        strAlpha = strAlpha + "B";
                        break;
                    case '9':
                        strAlpha = strAlpha + "C";
                        break;
                }
            }
            return strAlpha;
        }
        #endregion
        #region Calculate TAT Record
        public static void CalculateTAT(IOrganizationService service, Guid RecId)
        {
            msdyn_workorder iWrkOd = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, RecId, new ColumnSet("createdon", "hil_jobclosuredon", "hil_tattime"));
            msdyn_workorder iUpdate = new msdyn_workorder();
            iUpdate.Id = iWrkOd.Id;
            DateTime WorkDoneOn = new DateTime();
            DateTime CreatedOn = new DateTime();
            CreatedOn = Convert.ToDateTime(iWrkOd.CreatedOn);
            if (iWrkOd.Contains("hil_tattime"))
            {
                WorkDoneOn = Convert.ToDateTime(iWrkOd["hil_tattime"]);
                TimeSpan diff = WorkDoneOn - CreatedOn;
                double hours = diff.TotalMinutes / 60;
                iUpdate["hil_tattimecalculated"] = Convert.ToDecimal(hours);

                //iUpdate.hil_TimeinJobClosure = Convert.ToDecimal(hours);
                if (hours > 72)
                {
                    iUpdate.hil_SLAStatus = new OptionSetValue(1);
                }
                else
                {
                    iUpdate.hil_SLAStatus = new OptionSetValue(4);
                }
                #region  Added by Kuldeep Khare on 25/Nov/2019 to capture Job TAT Category
                EntityReference entRef = Havells_Plugin.WorkOrder.Common.TATCategory(service, hours);
                if (entRef != null)
                {
                    iUpdate["hil_tatcategory"] = entRef;
                }
                #endregion
                service.Update(iUpdate);
            }
        }
        #endregion
        #region Check If Actions Used
        public bool CheckIfActionsUsed(IOrganizationService service, Guid JobId)
        {
            bool IfNotOkay = true;
            QueryExpression Query = new QueryExpression(msdyn_workorderservice.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("msdyn_workorder", ConditionOperator.Equal, JobId);
            Query.Criteria.AddCondition("msdyn_linestatus", ConditionOperator.Equal, 690970001);
            EntityCollection Found1 = service.RetrieveMultiple(Query);
            if (Found1.Entities.Count > 0)
            {
                IfNotOkay = false;
            }
            return IfNotOkay;
        }
        #endregion
        #region Convert Number to words
        //public string ConvertNumberToWords(long number)
        //{
        //    if (number == 0) return "Zero";
        //    if (number < 0) return "minus " + ConvertNumbertoWords(Math.Abs(number));
        //    string words = "";
        //    if ((number / 100000) > 0)
        //    {
        //        words += ConvertNumbertoWords(number / 100000) + " lakh ";
        //        number %= 100000;
        //    }
        //    if ((number / 1000) > 0)
        //    {
        //        words += ConvertNumbertoWords(number / 1000) + " thousand ";
        //        number %= 1000;
        //    }
        //    if ((number / 100) > 0)
        //    {
        //        words += ConvertNumbertoWords(number / 100) + " hundred ";
        //        number %= 100;
        //    }
        //    if (number > 0)
        //    {
        //        if (words != "") words += "and ";
        //        var unitsMap = new[]
        //            {
        //            "zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen"
        //            };
        //        var tensMap = new[]
        //            {
        //            "zero", "ten", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"
        //            };
        //        if (number < 20) words += unitsMap[number];
        //        else
        //        {
        //            words += tensMap[number / 10];
        //            if ((number % 10) > 0) words += " " + unitsMap[number % 10];
        //        }
        //    }
        //    return words;
        //}
        //public static void UpdateReceiptNumberToWordOnTicket(IOrganizationService service, Guid TicketId)
        //{
        //    decimal ReceiptAmount = 0;
        //    msdyn_workorder Ticket = new msdyn_workorder();
        //    Ticket.Id = TicketId;
        //    QueryExpression Query = new QueryExpression(hil_estimate.EntityLogicalName);
        //    Query.ColumnSet = new ColumnSet("hil_receiptamount", "hil_modeofpayment");
        //    Query.Criteria = new FilterExpression(LogicalOperator.And);
        //    Query.Criteria.AddCondition("hil_job", ConditionOperator.Equal, TicketId);
        //    EntityCollection Found1 = service.RetrieveMultiple(Query);
        //}
        #endregion

        #region Create Child Job
        public static EntityReference GetManagerId(IOrganizationService service, Guid userGuId)
        {
            EntityReference _managerRecord = null;
            try
            {
                QueryExpression qeObj = new QueryExpression(SystemUser.EntityLogicalName);
                qeObj.ColumnSet = new ColumnSet("parentsystemuserid", "positionid", "systemuserid");
                qeObj.Criteria = new FilterExpression(LogicalOperator.And);
                qeObj.Criteria.AddCondition(new ConditionExpression("systemuserid", ConditionOperator.Equal, userGuId));
                EntityCollection ecObj = service.RetrieveMultiple(qeObj);
                if (ecObj.Entities.Count > 0)
                {
                    if (!ecObj.Entities[0].Contains("positionid"))
                    {
                        if (ecObj.Entities[0].Contains("parentsystemuserid"))
                        {
                            _managerRecord = ecObj.Entities[0].GetAttributeValue<EntityReference>("parentsystemuserid");
                        }
                    }
                    else if (ecObj.Entities[0].GetAttributeValue<EntityReference>("positionid").Name.ToUpper() != "BSH")
                    {
                        if (ecObj.Entities[0].Contains("parentsystemuserid"))
                        {
                            _managerRecord = GetManagerId(service, ecObj.Entities[0].GetAttributeValue<EntityReference>("parentsystemuserid").Id);
                        }
                    }
                    else
                    {
                        _managerRecord = ecObj.Entities[0].ToEntityReference();
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
            return _managerRecord;
        }
        //public static EntityReference GetManagerId(IOrganizationService service, Guid assigneeId)
        //{
        //    EntityReference _managerRecord = null;
        //    try
        //    {
        //        QueryExpression qeObj = new QueryExpression(SystemUser.EntityLogicalName);
        //        qeObj.ColumnSet = new ColumnSet("parentsystemuserid", "positionid");
        //        qeObj.Criteria = new FilterExpression(LogicalOperator.And);
        //        qeObj.Criteria.AddCondition(new ConditionExpression("systemuserid", ConditionOperator.Equal, assigneeId));
        //        EntityCollection ecObj = service.RetrieveMultiple(qeObj);
        //        if (ecObj.Entities.Count > 0)
        //        {
        //            foreach (SystemUser ent in ecObj.Entities)
        //            {
        //                if (!ent.Contains("positionid"))
        //                {
        //                    if (ent.Contains("parentsystemuserid"))
        //                    {
        //                        _managerRecord = ent.GetAttributeValue<EntityReference>("parentsystemuserid");
        //                    }
        //                }
        //                else
        //                {
        //                    if (ent.Contains("parentsystemuserid"))
        //                    {
        //                        if (ent.GetAttributeValue<EntityReference>("positionid").Name == "DSE")
        //                        { _managerRecord = ent.GetAttributeValue<EntityReference>("parentsystemuserid"); }
        //                        else if (ent.GetAttributeValue<EntityReference>("positionid").Name == "Franchise Technician")
        //                        { _managerRecord = GetManagerId(service, ent.GetAttributeValue<EntityReference>("parentsystemuserid").Id); }
        //                        else
        //                        { _managerRecord = ent.GetAttributeValue<EntityReference>("parentsystemuserid"); }
        //                    }
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new InvalidPluginExecutionException("Havells_Plugin.WorkOrder.PostUpdate.Execute.GetManagerId: " + ex.Message);
        //    }
        //    return _managerRecord;
        //}
        public static void CreateChildJob(IOrganizationService service, msdyn_workorder enWorkorder)
        {
            try
            {
                Guid _childWorkOrderId = Guid.Empty;
                msdyn_workorder enWorkorderUpdate = new msdyn_workorder();
                tracingService.Trace("K1");
                Guid PriceList = Helper.GetGuidbyName(PriceLevel.EntityLogicalName, "name", "Default Price List", service);
                Guid IncidentType = Helper.GetGuidbyName(msdyn_incidenttype.EntityLogicalName, "msdyn_name", "Default Value", service);
                Guid SrvAccount = Helper.GetGuidbyName(Account.EntityLogicalName, "name", "Dummy Customer", service);

                msdyn_workorder enChildWorkorder = new msdyn_workorder();
                if (enWorkorder.hil_CustomerRef != null)
                { enChildWorkorder.hil_CustomerRef = enWorkorder.hil_CustomerRef; }
                if (enWorkorder.hil_customername != null)
                { enChildWorkorder.hil_customername = enWorkorder.hil_customername; }
                if (enWorkorder.hil_mobilenumber != null)
                { enChildWorkorder.hil_mobilenumber = enWorkorder.hil_mobilenumber; }
                if (enWorkorder.hil_CallingNumber != null)
                { enChildWorkorder.hil_CallingNumber = enWorkorder.hil_CallingNumber; }
                if (enWorkorder.hil_Alternate != null)
                { enChildWorkorder.hil_Alternate = enWorkorder.hil_Alternate; }
                if (enWorkorder.hil_Address != null)
                { enChildWorkorder.hil_Address = enWorkorder.hil_Address; }
                if (enWorkorder.hil_FullAddress != null)
                { enChildWorkorder.hil_FullAddress = enWorkorder.hil_FullAddress; }
                if (enWorkorder.hil_pincode != null)
                { enChildWorkorder.hil_pincode = enWorkorder.hil_pincode; }
                if (enWorkorder.hil_PinCodeText != null)
                { enChildWorkorder.hil_PinCodeText = enWorkorder.hil_PinCodeText; }
                if (enWorkorder.hil_Region != null)
                { enChildWorkorder.hil_Region = enWorkorder.hil_Region; }
                if (enWorkorder.hil_RegionText != null)
                { enChildWorkorder.hil_RegionText = enWorkorder.hil_RegionText; }
                if (enWorkorder.hil_Branch != null)
                { enChildWorkorder.hil_Branch = enWorkorder.hil_Branch; }
                if (enWorkorder.hil_BranchText != null)
                { enChildWorkorder.hil_BranchText = enWorkorder.hil_BranchText; }
                if (enWorkorder.hil_City != null)
                { enChildWorkorder.hil_City = enWorkorder.hil_City; }
                if (enWorkorder.hil_CityText != null)
                { enChildWorkorder.hil_CityText = enWorkorder.hil_CityText; }
                if (enWorkorder.hil_district != null)
                { enChildWorkorder.hil_district = enWorkorder.hil_district; }
                if (enWorkorder.hil_DistrictText != null)
                { enChildWorkorder.hil_DistrictText = enWorkorder.hil_DistrictText; }
                if (enWorkorder.hil_state != null)
                { enChildWorkorder.hil_state = enWorkorder.hil_state; }
                if (enWorkorder.hil_StateText != null)
                { enChildWorkorder.hil_StateText = enWorkorder.hil_StateText; }
                if (enWorkorder.GetAttributeValue<EntityReference>("hil_salesoffice") != null)
                { enChildWorkorder["hil_salesoffice"] = enWorkorder.GetAttributeValue<EntityReference>("hil_salesoffice"); }
                if (enWorkorder.hil_Email != null)
                { enChildWorkorder.hil_Email = enWorkorder.hil_Email; }
                if (enWorkorder.hil_Productcategory != null)
                { enChildWorkorder.hil_Productcategory = enWorkorder.hil_Productcategory; }
                if (enWorkorder.hil_ProductSubcategory != null)
                { enChildWorkorder.hil_ProductSubcategory = enWorkorder.hil_ProductSubcategory; }
                if (enWorkorder.hil_ProductCatSubCatMapping != null)
                { enChildWorkorder.hil_ProductCatSubCatMapping = enWorkorder.hil_ProductCatSubCatMapping; }
                if (enWorkorder.GetAttributeValue<EntityReference>("hil_consumertype") != null)
                { enChildWorkorder["hil_consumertype"] = enWorkorder.GetAttributeValue<EntityReference>("hil_consumertype"); }
                if (enWorkorder.GetAttributeValue<EntityReference>("hil_consumercategory") != null)
                { enChildWorkorder["hil_consumercategory"] = enWorkorder.GetAttributeValue<EntityReference>("hil_consumercategory"); }
                if (enWorkorder.hil_natureofcomplaint != null)
                { enChildWorkorder.hil_natureofcomplaint = enWorkorder.hil_natureofcomplaint; }
                if (enWorkorder.hil_CallSubType != null)
                { enChildWorkorder.hil_CallSubType = enWorkorder.hil_CallSubType; }
                if (enWorkorder.hil_CallType != null)
                { enChildWorkorder.hil_CallType = enWorkorder.hil_CallType; }
                if (enWorkorder.hil_quantity != null)
                { enChildWorkorder.hil_quantity = enWorkorder.hil_quantity; }
                if (enWorkorder.GetAttributeValue<OptionSetValue>("hil_callertype") != null)
                { enChildWorkorder["hil_callertype"] = enWorkorder.GetAttributeValue<OptionSetValue>("hil_callertype"); }
                if (enWorkorder.hil_SourceofJob != null)
                { enChildWorkorder.hil_SourceofJob = enWorkorder.hil_SourceofJob; }
                enChildWorkorder.msdyn_ServiceAccount = new EntityReference(Account.EntityLogicalName, SrvAccount);
                enChildWorkorder.msdyn_BillingAccount = new EntityReference(Account.EntityLogicalName, SrvAccount);
                if (IncidentType != Guid.Empty)
                { enChildWorkorder.msdyn_PrimaryIncidentType = new EntityReference(msdyn_incidenttype.EntityLogicalName, IncidentType); }
                if (PriceList != Guid.Empty)
                { enChildWorkorder.msdyn_PriceList = new EntityReference(PriceLevel.EntityLogicalName, PriceList); }
                if (enWorkorder.msdyn_CustomerAsset != null)
                { enChildWorkorder.msdyn_CustomerAsset = enWorkorder.msdyn_CustomerAsset; }
                if (enWorkorder.hil_PurchaseDate != null)
                { enChildWorkorder.hil_PurchaseDate = enWorkorder.hil_PurchaseDate; }
                enChildWorkorder.msdyn_ParentWorkOrder = new EntityReference("msdyn_workorder", enWorkorder.Id);
                enChildWorkorder["hil_jobreopenreason"] = new OptionSetValue(910590001); // {"KKG Audit Failure"}

                _childWorkOrderId = service.Create(enChildWorkorder);
                tracingService.Trace("K2");
                if (_childWorkOrderId != Guid.Empty)
                {
                    tracingService.Trace("K3");
                    enWorkorderUpdate["hil_newjobid"] = new EntityReference("msdyn_workorder", _childWorkOrderId);
                    EntityReference er = GetManagerId(service, enWorkorder.OwnerId.Id);
                    if (er != null)
                    {
                        enWorkorderUpdate["hil_bsh"] = er;
                    }
                    enWorkorderUpdate.Id = enWorkorder.Id;
                    service.Update(enWorkorderUpdate);
                    CreateKKGAuditFailedSAWActivity(enWorkorder.Id, service);
                }
                tracingService.Trace("K4");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        #endregion

        #region Update Warranty Date on Job
        public static void UpdateAssetAndWarrantyDetailsOnJob(IOrganizationService service, Entity entity)
        {
            try
            {
                msdyn_workorder iWo = entity.ToEntity<msdyn_workorder>();
                msdyn_workorderincident iHeader = new msdyn_workorderincident();

                QueryExpression QueryPR = new QueryExpression(msdyn_workorderincident.EntityLogicalName);
                QueryPR.ColumnSet = new ColumnSet("new_warrantyenddate", "hil_warrantystatus", "msdyn_customerasset", "hil_modelname", "msdyn_workorderincidentid");
                QueryPR.Criteria = new FilterExpression(LogicalOperator.And);
                QueryPR.Criteria.AddCondition(new ConditionExpression("msdyn_workorder", ConditionOperator.Equal, entity.Id));
                QueryPR.Criteria.AddCondition(new ConditionExpression("statuscode", ConditionOperator.Equal, 1));
                QueryPR.TopCount = 1;
                QueryPR.AddOrder("createdon", OrderType.Ascending);
                EntityCollection Found = service.RetrieveMultiple(QueryPR);
                if (Found.Entities.Count > 0)
                {
                    foreach (msdyn_workorderincident iRec in Found.Entities)
                    {
                        if (iRec.Attributes.Contains("new_warrantyenddate"))
                        {
                            iWo["hil_warrantydate"] = iRec.GetAttributeValue<DateTime>("new_warrantyenddate");
                            iWo["hil_warrantystatus"] = iRec.GetAttributeValue<OptionSetValue>("hil_warrantystatus");
                            iWo["msdyn_customerasset"] = iRec.GetAttributeValue<EntityReference>("msdyn_customerasset");
                            iWo["hil_modelname"] = iRec.GetAttributeValue<string>("hil_modelname");
                            service.Update(iWo);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        #endregion
    }
}