using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.CRM.Plugin.Approval
{
    enum ApprovalPurpose
    {
        Tender_Fee_Approval = 1,
        BG_Request_Approval = 2,
        BG_Amendment_Approval = 3,
        Order_Price_Approval = 4
    }
    enum EntityObjectType
    {
        Tender_Enquiry = 1,
        Order_Check_List = 2,
        Bank_Guarantee = 3
    }
    enum ApprovalStatus
    {
        Approved = 1,
        Rejected = 2,
        Submitted_for_Approval = 3,
        Draft = 4,
        Level_1_Approved = 5,
        Level_2_Approved = 6,
        Level_3_Approved = 7,
        Level_4_Approved = 8,
        Level_5_Approved = 9,
        Partial_Approved = 10,
        Send_for_Revision = 11
    }
    public class HelperClass
    {
        public static void CreateApprovals(IOrganizationService service, ITracingService tracingService, string entityName, string entityId, int purpose, int entityObjectType)
        {
            if (purpose == (int)ApprovalPurpose.Order_Price_Approval && entityObjectType == (int)EntityObjectType.Order_Check_List)
            {
                CreateOCLPriceApproval(service, null, entityName, entityId, purpose, entityObjectType);
            }
            else
            {
                #region variables
                EntityReference _departmentRef = null;
                EntityReference _salesOfficeRef = null;
                EntityReference _UserHerirchy = null;
                string currentDate = DateTime.Now.ToString("yyyy-MM-dd");

                #endregion
                tracingService.Trace("d");
                Entity targetEntity = service.Retrieve(entityName, new Guid(entityId), new ColumnSet(true));
                if (targetEntity.Contains("hil_department"))
                {
                    _departmentRef = (EntityReference)targetEntity["hil_department"];
                }
                if (targetEntity.Contains("hil_salesoffice"))
                {
                    _salesOfficeRef = (EntityReference)targetEntity["hil_salesoffice"];
                }
                if (targetEntity.Contains("hil_onbehalfuser"))
                {
                    _UserHerirchy = (EntityReference)targetEntity["hil_onbehalfuser"];
                }
                EntityCollection approvalLines = RetriveApprovalsMatrixLine(service, currentDate, _departmentRef.Id, _salesOfficeRef.Id, purpose, entityObjectType);
                
                createApprovals(service, approvalLines, targetEntity, null, _UserHerirchy);
            }
        }
        public static void CreateOCLPriceApproval(IOrganizationService service, ITracingService tracingService, string entityName, string entityId, int purpose, int entityObjectType)
        {
            #region variables
            EntityReference _departmentRef = null;
            EntityReference _salesOfficeRef = null;
            EntityReference _UserHerirchy = null;

            #endregion

            Entity targetEntity = service.Retrieve(entityName, new Guid(entityId), new ColumnSet(true));
            if (targetEntity.Contains("hil_department"))
            {
                _departmentRef = (EntityReference)targetEntity["hil_department"];
            }
            if (targetEntity.Contains("hil_salesoffice"))
            {
                _salesOfficeRef = (EntityReference)targetEntity["hil_salesoffice"];
            }
            if (targetEntity.Contains("hil_onbehalfuser"))
            {
                _UserHerirchy = (EntityReference)targetEntity["hil_onbehalfuser"];
            }

            #region Retriev Max Discount MG Wise of all OCL Product
            EntityCollection oclProductColl = null;
            //Retriev all OCL Product
            string fetchPODiscount = @"<fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
	                            <entity name='hil_orderchecklistproduct'>
		                            <attribute name='hil_materialgroup' alias='hil_materialgroup' groupby='true' />
		                            <attribute name='hil_podiscount' alias='amount' aggregate='max' />
		                            <filter type='and'>
			                            <condition attribute='hil_orderchecklistid' operator='eq' value='" + targetEntity.Id + @"' />
			                            <condition attribute='statecode' operator='eq' value='0' />
		                            </filter>
	                            </entity>
                            </fetch>";
            string fetchOrderBookingDiscount = @"<fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
	                            <entity name='hil_orderchecklistproduct'>
		                            <attribute name='hil_materialgroup' alias='hil_materialgroup' groupby='true' />
		                            <attribute name='hil_discount' alias='amount' aggregate='max' />
		                            <filter type='and'>
			                            <condition attribute='hil_orderchecklistid' operator='eq' value='" + targetEntity.Id + @"' />
			                            <condition attribute='statecode' operator='eq' value='0' />
		                            </filter>
	                            </entity>
                            </fetch>";
            if (targetEntity.Contains("hil_tenderno"))
            {
                oclProductColl = service.RetrieveMultiple(new FetchExpression(fetchPODiscount));
                tracingService.Trace("d");
            }
            else
            {
                oclProductColl = service.RetrieveMultiple(new FetchExpression(fetchOrderBookingDiscount));
                tracingService.Trace("d");
            }
            #endregion
            List<ApprovalsClass> approvalLines = new List<ApprovalsClass>();
            foreach (Entity ent in oclProductColl.Entities)
            {
                string currentDate = DateTime.Now.ToString("yyyy-MM-dd");
                decimal maxDiscount = (decimal)ent.GetAttributeValue<AliasedValue>("amount").Value;
                EntityReference _materialGroupRef = (EntityReference)ent.GetAttributeValue<AliasedValue>("hil_materialgroup").Value;
                approvalLines.AddRange(RetriveApprovalsMatrixLineforPrice(service, currentDate, maxDiscount, _materialGroupRef.Id, _departmentRef.Id, _salesOfficeRef.Id, purpose, entityObjectType, _UserHerirchy));
                tracingService.Trace("d");
            }
            //Console.Clear();
            approvalLines = approvalLines.OrderBy(x => x.ApproverID).ThenBy(x => x.Level).ToList();
            tracingService.Trace("******************************************************");
            foreach (ApprovalsClass x in approvalLines)
            {
                tracingService.Trace("ApproverName " + x.ApprovalName);
                tracingService.Trace(" | ApproverID " + x.ApproverID);
                tracingService.Trace(" | Level " + x.Level);
                tracingService.Trace(" | Material Group ID " + x.MAterialGroupName);
                tracingService.Trace(" | Material Group Name " + x.MAterialGroupName);
            }

            List<Guid> approverList = approvalLines.Select(x => x.ApproverID).Distinct().ToList();
            List<ApprovalsClass>[] approvalLines123 = new List<ApprovalsClass>[approverList.Count];
            int i = 0;
            foreach (Guid guid in approverList)
            {
                approvalLines123[i] = approvalLines.Where(x => x.ApproverID == guid).ToList();
                i++;
            }
            createApprovalsforOCLPrice(service, approvalLines123, targetEntity, tracingService);

            submitForApproval(service, targetEntity.ToEntityReference(), approverList, tracingService);
            tracingService.Trace("dd");
            //createApprovals(service, approvalLines, targetEntity, null, _UserHerirchy);
        }
        public static List<ApprovalsClass> RetriveApprovalsMatrixLineforPrice(IOrganizationService service, string currentDate, decimal maxDiscount, Guid _MatrialGroupId, Guid _DepartmentId, Guid _SalesOfficeId, int purpose, int entityObjectType, EntityReference _UserHerirchy)
        {
            string dis = maxDiscount.ToString().Split('.')[0];
            string fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='hil_approvalmatrixline'>
                                    <attribute name='hil_name' />
                                    <attribute name='hil_maxdiscountpercentage' />
                                    <attribute name='hil_level' />
                                    <attribute name='hil_effectiveto' />
                                    <attribute name='hil_effectivefrom' />
                                    <attribute name='hil_approvalposition' />
                                    <attribute name='hil_approvalmatrixheader' />
                                    <attribute name='hil_approval' />
                                    <attribute name='hil_isdynamicapproval' />
                                    <attribute name='statecode' />
                                    <attribute name='hil_materialgroup' />
                                    <attribute name='hil_approvalmatrixlineid' />
                                    <order attribute='hil_name' descending='false' />
                                    <filter type='and'>
                                      <condition attribute='hil_materialgroup' operator='eq' value='{{{_MatrialGroupId}}}' />
                                      <condition attribute='hil_effectivefrom' operator='on-or-before' value='{currentDate}' />
                                      <condition attribute='hil_effectiveto' operator='on-or-after' value='{currentDate}' />
                                      <condition attribute='hil_maxdiscountpercentage' operator='le' value='{dis}' />
                                      <condition attribute='statecode' operator='eq' value='0' />
                                    </filter>
                                    <link-entity name='hil_approvalmatrix' from='hil_approvalmatrixid' to='hil_approvalmatrixheader'
                                      link-type='inner' alias='aa'>
                                      <filter type='and'>
                                        <condition attribute='hil_department' operator='eq' value='{{{_DepartmentId}}}' />
                                        <condition attribute='hil_entityobjecttype' operator='eq' value='{entityObjectType}' />
                                        <condition attribute='hil_approvalpurpose' operator='eq' value='{purpose}' />
                                        <condition attribute='hil_salesoffice' operator='eq' value='{{{_SalesOfficeId}}}' />
                                        <condition attribute='statecode' operator='eq' value='0' />
                                      </filter>
                                    </link-entity>
                                  </entity>
                                </fetch>";
            EntityCollection entityCollection = service.RetrieveMultiple(new FetchExpression(fetchXML));
            List<ApprovalsClass> approvalsClasses = new List<ApprovalsClass>();
            EntityReference approver = new EntityReference();
            EntityReference MAterialGroup = new EntityReference();
            foreach (Entity entity in entityCollection.Entities)
            {
                approver = null;
                MAterialGroup = null;
                ApprovalsClass approvalsClass = new ApprovalsClass();
                approvalsClass.Level = entity.GetAttributeValue<OptionSetValue>("hil_level").Value;
                if ((bool)entity["hil_isdynamicapproval"])
                {
                    String approvalPosition = entity.GetAttributeValue<EntityReference>("hil_approvalposition").Name;
                    if (_UserHerirchy != null)
                        approver = getApproverByPosition(_UserHerirchy, approvalPosition, service, null);
                    else
                        throw new InvalidPluginExecutionException("Approver not Found");
                }
                else
                {
                    approver = entity.GetAttributeValue<EntityReference>("hil_approval");
                }
                if (approver != null)
                {
                    approvalsClass.ApproverID = approver.Id;
                    approvalsClass.ApprovalName = approver.Name;
                    approvalsClass.ApprovalEntity = approver.LogicalName;
                }
                if (entity.Contains("hil_materialgroup"))
                {
                    MAterialGroup = entity.GetAttributeValue<EntityReference>("hil_materialgroup");

                    approvalsClass.MAterialGroupID = MAterialGroup.Id;
                    approvalsClass.MAterialGroupName = MAterialGroup.Name;
                    approvalsClass.MAterialGroupEntity = MAterialGroup.LogicalName;
                }
                approvalsClass.ApprovalMatrixLineEntity = entity;
                approvalsClasses.Add(approvalsClass);
            }
            return approvalsClasses;
        }
        public static EntityCollection RetriveApprovalsMatrixLine(IOrganizationService service, string currentDate, Guid _DepartmentId, Guid _SalesOfficeId, int purpose, int entityObjectType)
        {
            string fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='hil_approvalmatrixline'>
                                    <attribute name='hil_name' />
                                    <attribute name='hil_maxdiscountpercentage' />
                                    <attribute name='hil_level' />
                                    <attribute name='hil_effectiveto' />
                                    <attribute name='hil_effectivefrom' />
                                    <attribute name='hil_approvalposition' />
                                    <attribute name='hil_approvalmatrixheader' />
                                    <attribute name='hil_approval' />
                                    <attribute name='hil_isdynamicapproval' />
                                    <attribute name='statecode' />
                                    <attribute name='hil_approvalmatrixlineid' />
                                    <order attribute='hil_name' descending='false' />
                                    <filter type='and'>
                                      <condition attribute='hil_effectivefrom' operator='on-or-before' value='{currentDate}' />
                                      <condition attribute='hil_effectiveto' operator='on-or-after' value='{currentDate}' />
                                      <condition attribute='statecode' operator='eq' value='0' />
                                    </filter>
                                    <link-entity name='hil_approvalmatrix' from='hil_approvalmatrixid' to='hil_approvalmatrixheader'
                                      link-type='inner' alias='aa'>
                                      <filter type='and'>
                                        <condition attribute='hil_department' operator='eq' value='{_DepartmentId}' />
                                        <condition attribute='hil_entityobjecttype' operator='eq' value='{entityObjectType}' />
                                        <condition attribute='hil_approvalpurpose' operator='eq' value='{purpose}' />
                                        <condition attribute='hil_salesoffice' operator='eq' value='{_SalesOfficeId}' />
                                        <condition attribute='statecode' operator='eq' value='0' />
                                      </filter>
                                    </link-entity>
                                  </entity>
                                </fetch>";
            return service.RetrieveMultiple(new FetchExpression(fetchXML));
        }
        public static void createApprovals(IOrganizationService service, EntityCollection ApprovalsMatrixLines, Entity target, EntityReference MGref, EntityReference _UserHerirchy)
        {
            Guid lastApproval = Guid.Empty;
            Guid currentApproval = Guid.Empty;
            for (int i = 0; i < ApprovalsMatrixLines.Entities.Count; i++)
            {
                Entity approvalMatrixLine = ApprovalsMatrixLines[i];
                EntityReference targetOwner = target.GetAttributeValue<EntityReference>("ownerid");

                Entity _approval = new Entity("hil_approval");
                EntityReference approver = new EntityReference();
                _approval["subject"] = approvalMatrixLine.GetAttributeValue<string>("hil_name") + "_Level " + ApprovalsMatrixLines[i].GetAttributeValue<OptionSetValue>("hil_level").Value;
                //tracingService.Trace("levrl");
                if ((bool)approvalMatrixLine["hil_isdynamicapproval"])
                {
                    String approvalPosition = approvalMatrixLine.GetAttributeValue<EntityReference>("hil_approvalposition").Name;
                    if (_UserHerirchy != null)
                        approver = getApproverByPosition(_UserHerirchy, approvalPosition, service, null);
                    else
                        throw new InvalidPluginExecutionException("Approver not Found");
                }
                else
                {
                    approver = approvalMatrixLine.GetAttributeValue<EntityReference>("hil_approval");
                }
                _approval["ownerid"] = approver;
                _approval["hil_approvaltempate"] = approvalMatrixLine.ToEntityReference();
                _approval["hil_level"] = approvalMatrixLine.GetAttributeValue<OptionSetValue>("hil_level");
                _approval["regardingobjectid"] = target.ToEntityReference();// new EntityReference(entityName, new Guid(entityId));
                if (approvalMatrixLine.GetAttributeValue<OptionSetValue>("hil_level").Value == 1)
                {
                    _approval["hil_approvalstatus"] = new OptionSetValue(3);// 3 for - submit for approval    4 for - Draft
                    _approval["hil_requesteddate"] = DateTime.Now;
                }
                else
                {
                    _approval["hil_approvalstatus"] = new OptionSetValue(4);// 4 for - Draft
                }

                currentApproval = service.Create(_approval);
                if (approvalMatrixLine.GetAttributeValue<OptionSetValue>("hil_level").Value != 1)
                {
                    Entity entity = new Entity("hil_approval");
                    entity.Id = lastApproval;
                    entity["hil_nextapproval"] = new EntityReference(_approval.LogicalName, currentApproval);
                    service.Update(entity);
                    lastApproval = currentApproval;
                }
                else
                {
                    lastApproval = currentApproval;
                }
                if (target.LogicalName == "hil_orderchecklist")
                    CreateOCLPrductApproval(service, target.ToEntityReference(), MGref, new EntityReference("hil_approval", currentApproval), (OptionSetValue)_approval["hil_approvalstatus"]);
            }
        }
        public static void createApprovalsforOCLPrice(IOrganizationService service, List<ApprovalsClass>[] approvalHeadrs, Entity target,ITracingService tracingService)
        {
            EntityReference approvalHeader = null;
            foreach (List<ApprovalsClass> approvalsClass in approvalHeadrs)
            {
                Entity approvalMatrixLine = approvalsClass[0].ApprovalMatrixLineEntity as Entity;
                Entity _approval = new Entity("hil_approval");
                EntityReference approver = new EntityReference();
                string MGList = string.Empty;
                for (int i = 0; i < approvalsClass.Count; i++)
                {
                    if (i == approvalsClass.Count - 2)
                        MGList = MGList + approvalsClass[i].MAterialGroupName + " and ";
                    else
                        MGList = MGList + approvalsClass[i].MAterialGroupName + ", ";
                }
                //foreach (ApprovalsClass approvalsClass1 in approvalsClass)
                //{
                //    MGList = MGList + approvalsClass1.MAterialGroupName + ", ";
                //}
                _approval["subject"] = "Approvals for MG : " + MGList.Substring(0, MGList.Length - 2);
                approver = new EntityReference(approvalsClass[0].ApprovalEntity, approvalsClass[0].ApproverID);
                _approval["ownerid"] = approver;
                _approval["regardingobjectid"] = target.ToEntityReference();
                approvalHeader = new EntityReference("hil_approval", service.Create(_approval));
                tracingService.Trace("Approval Header Created..");
                createApprovalsLinesforOCLPrice(service, approvalsClass, target, approvalHeader, ("Approvals for MG : " + MGList.Substring(0, MGList.Length - 2)), approver,tracingService);
            }
        }
        public static void createApprovalsLinesforOCLPrice(IOrganizationService service, List<ApprovalsClass> approvalHeadrs, Entity target,
            EntityReference approvalHeader, string headerName, EntityReference approver, ITracingService tracingService)
        {
            foreach (ApprovalsClass approvalsClass in approvalHeadrs)
            {
                Entity approvalMatrixLine = approvalsClass.ApprovalMatrixLineEntity as Entity;

                string fetchXML = $@"<fetch version=""1.0"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"">
                                      <entity name=""hil_orderchecklistproduct"">
                                        <attribute name=""hil_orderchecklistproductid"" />
                                        <attribute name=""hil_name"" />
                                        <attribute name=""createdon"" />
                                        <order attribute=""hil_name"" descending=""false"" />
                                        <filter type=""and"">
                                          <condition attribute=""hil_orderchecklistid"" operator=""eq"" value=""{target.Id}"" />
                                          <condition attribute=""hil_materialgroup"" operator=""eq"" value=""{approvalsClass.MAterialGroupID}"" />
                                          <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                                        </filter>
                                      </entity>
                                    </fetch>";
                EntityCollection oclPrd = service.RetrieveMultiple(new FetchExpression(fetchXML));
                foreach (Entity entity in oclPrd.Entities)
                {
                    Entity _approval = new Entity("hil_oclproductapproval");
                    _approval["hil_approval"] = approvalHeader;
                    _approval["hil_name"] = headerName + "_item_" + entity.GetAttributeValue<string>("hil_name");
                    _approval["hil_approvalstatus"] = new OptionSetValue(4);
                    _approval["ownerid"] = approver;
                    _approval["hil_oclproduct"] = entity.ToEntityReference();
                    _approval["hil_level"] = approvalMatrixLine.GetAttributeValue<OptionSetValue>("hil_level");
                    service.Create(_approval);
                    tracingService.Trace("Approval Line Created..");
                }
            }
        }
        public static void submitForApproval(IOrganizationService service, EntityReference target, List<Guid> approvalList, ITracingService tracingService)
        {
            tracingService.Trace("**********************");
            bool sendForApproval = true;
            foreach (Guid approverId in approvalList)
            {
                sendForApproval = true;
                int approvalStatus = (int)ApprovalStatus.Draft;
                string fetchXML = $@"<fetch version=""1.0"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"">
                                      <entity name=""hil_oclproductapproval"">
                                        <attribute name=""hil_level"" />
                                        <attribute name=""hil_approvalstatus"" />
                                        <attribute name=""hil_approval"" />
                                        <attribute name=""hil_oclproductapprovalid"" />
                                        <attribute name=""ownerid"" />
                                        <attribute name=""hil_oclproduct"" />
                                        <order attribute=""hil_level"" descending=""true"" />
                                        <filter type=""and"">
                                          <condition attribute=""hil_approvalstatus"" operator=""eq"" value=""{approvalStatus}"" />
                                          <condition attribute=""ownerid"" operator=""eq"" value=""{approverId}"" />
                                        </filter>
                                        <link-entity name=""hil_approval"" from=""activityid"" to=""hil_approval"" link-type=""inner"" alias=""ocl"">
                                          <attribute name=""ownerid"" />
                                          <attribute name=""hil_approvaldate"" />
                                          <filter type=""and"">
                                            <condition attribute=""regardingobjectid"" operator=""eq"" value=""{target.Id}"" />
                                          </filter>
                                        </link-entity>
                                      </entity>
                                    </fetch>";
                EntityCollection entityCollection = service.RetrieveMultiple(new FetchExpression(fetchXML));
                if (entityCollection.Entities.Count == 0)
                    sendForApproval = false;
                if (sendForApproval)
                    foreach (Entity entity in entityCollection.Entities)
                    {
                        int level = entity.GetAttributeValue<OptionSetValue>("hil_level").Value;
                        EntityReference oclPrd = entity.GetAttributeValue<EntityReference>("hil_oclproduct");
                        for (int x = level; x > 0; x--)
                        {
                            fetchXML = $@"<fetch version=""1.0"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"">
                                      <entity name=""hil_oclproductapproval"">
                                        <attribute name=""hil_level"" />
                                        <attribute name=""hil_approvalstatus"" />
                                        <attribute name=""hil_oclproductapprovalid"" />
                                        <attribute name=""ownerid"" />
                                        <attribute name=""hil_oclproduct"" />
                                        <order attribute=""hil_level"" descending=""true"" />
                                        <filter type=""and"">
                                          <condition attribute=""hil_level"" operator=""eq"" value=""{x}"" />
                                          <condition attribute=""hil_oclproduct"" operator=""eq"" value=""{oclPrd.Id}"" />
                                        </filter>
                                      </entity>
                                    </fetch>";
                            EntityCollection newCol = service.RetrieveMultiple(new FetchExpression(fetchXML));
                            foreach (Entity c in newCol.Entities)
                                if (
                                    c.GetAttributeValue<EntityReference>("ownerid").Id != approverId &&
                                        (
                                            approvalStatus != (int)ApprovalStatus.Approved &&
                                            approvalStatus != (int)ApprovalStatus.Rejected &&
                                            approvalStatus != (int)ApprovalStatus.Draft
                                        )
                                    )
                                    sendForApproval = false;
                            //if (newCol.Entities.Count == 0)
                            //    sendForApproval = false;
                        }
                    }
                tracingService.Trace("check for Pendincinmg");
                if (sendForApproval)
                {
                    EntityReference approvalheader = entityCollection[0].GetAttributeValue<EntityReference>("hil_approval");
                    Entity aplHead = new Entity(approvalheader.LogicalName, approvalheader.Id);
                    aplHead["hil_approvalstatus"] = new OptionSetValue((int)ApprovalStatus.Submitted_for_Approval);
                    service.Update(aplHead);
                    QueryExpression _query = new QueryExpression("hil_oclproductapproval");
                    _query.ColumnSet = new ColumnSet(false);
                    _query.Criteria = new FilterExpression(LogicalOperator.And);
                    _query.Criteria.AddCondition(new ConditionExpression("hil_approval", ConditionOperator.Equal, approvalheader.Id));
                    EntityCollection dd = service.RetrieveMultiple(_query);
                    foreach (Entity entity in dd.Entities)
                    {
                        entity["hil_approvalstatus"] = new OptionSetValue((int)ApprovalStatus.Submitted_for_Approval);
                        entity["hil_requestedon"] = DateTime.Now;
                        service.Update(entity);
                        tracingService.Trace("line Status changed");
                    }
                    tracingService.Trace("header Status changed");
                }
            }
        }
        public static void CreateOCLPrductApproval(IOrganizationService service, EntityReference OCLRef, EntityReference MGRef, EntityReference approval, OptionSetValue approvalStatus)
        {
            QueryExpression _query = new QueryExpression("hil_orderchecklistproduct");
            _query.ColumnSet = new ColumnSet(false);
            _query.Criteria = new FilterExpression(LogicalOperator.And);
            _query.Criteria.AddCondition(new ConditionExpression("hil_orderchecklistid", ConditionOperator.Equal, OCLRef.Id));
            _query.Criteria.AddCondition(new ConditionExpression("hil_materialgroup", ConditionOperator.Equal, MGRef.Id));
            _query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));

            EntityCollection oclPRdColl = service.RetrieveMultiple(_query);
            foreach (Entity oclPRd in oclPRdColl.Entities)
            {
                Entity OCLPrdApproval = new Entity("hil_oclproductapproval");
                OCLPrdApproval["hil_oclproduct"] = oclPRd.ToEntityReference();
                OCLPrdApproval["hil_approval"] = approval;
                OCLPrdApproval["hil_approvalstatus"] = approvalStatus;
                service.Create(OCLPrdApproval);

                Entity oclEntitiy = new Entity(oclPRd.LogicalName, oclPRd.Id);
                oclEntitiy["hil_approvalstatus"] = new OptionSetValue(3);
                service.Update(oclEntitiy);
            }
        }
        public static EntityReference getApproverByPosition(EntityReference userMapingRef, string approvalPosition, IOrganizationService service, ITracingService tracingService)
        {
            Entity userMaping = service.Retrieve(userMapingRef.LogicalName, userMapingRef.Id, new ColumnSet(true));
            //tracingService.Trace("getApproverByPosition Stratrd");
            EntityReference approvr = new EntityReference();
            if (approvalPosition == "Branch Product Head")
            {
                approvr = userMaping.GetAttributeValue<EntityReference>("hil_branchproducthead");
            }
            else if (approvalPosition == "Enquiry Creator")
            {
                approvr = userMaping.GetAttributeValue<EntityReference>("hil_user");
            }
            else if (approvalPosition == "Zonal Head")
            {
                approvr = userMaping.GetAttributeValue<EntityReference>("hil_zonalhead");
            }
            else if (approvalPosition == "BU Head")
            {
                approvr = userMaping.GetAttributeValue<EntityReference>("hil_buhead");
            }
            else if (approvalPosition == "Sr. Manager Finance")
            {
                QueryExpression query = new QueryExpression("systemuser");
                query.ColumnSet = new ColumnSet(true);
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition("positionid", ConditionOperator.Equal, new Guid("6532dc30-40a6-eb11-9442-6045bd72b6fd"));
                EntityCollection _entitys = service.RetrieveMultiple(query);

                if (_entitys.Entities.Count > 0)
                    approvr = _entitys[0].ToEntityReference();
            }
            else if (approvalPosition == "Sr. Manager Treasury")
            {
                QueryExpression query = new QueryExpression("user");
                query.ColumnSet = new ColumnSet(true);
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition("positionid", ConditionOperator.Equal, new Guid("{45C29172-3CD0-EB11-BACC-6045BD72E9C2}"));
                EntityCollection _entitys = service.RetrieveMultiple(query);
                if (_entitys.Entities.Count > 0)
                    approvr = _entitys[0].ToEntityReference();
            }
            else if (approvalPosition == "SCM")
            {
                approvr = userMaping.GetAttributeValue<EntityReference>("hil_scm");
            }
            //else if (approvalPosition == "Design Head")
            //{
            //    QueryExpression query = new QueryExpression("systemuser");
            //    query.ColumnSet = new ColumnSet(true);
            //    query.Criteria = new FilterExpression(LogicalOperator.And);
            //    query.Criteria.AddCondition("positionid", ConditionOperator.Equal, new Guid("6532dc30-40a6-eb11-9442-6045bd72b6fd"));
            //    EntityCollection _entitys = service.RetrieveMultiple(query);

            //    _approval["ownerid"] = entColl[i].GetAttributeValue<EntityReference>("hil_user"); = _entitys[0].ToEntityReference();
            //}
            //tracingService.Trace("getApproverByPosition end");
            return approvr;
        }
        public static void DeleteApprovals(IOrganizationService service, string oclID)
        {

            QueryExpression _query = new QueryExpression("hil_approval");
            _query.ColumnSet = new ColumnSet(false);
            _query.Criteria = new FilterExpression(LogicalOperator.And);
            _query.Criteria.AddCondition(new ConditionExpression("regardingobjectid", ConditionOperator.Equal, new Guid(oclID)));
            EntityCollection dd = service.RetrieveMultiple(_query);
            foreach (Entity entity in dd.Entities)
            {
                service.Delete(entity.LogicalName, entity.Id);
            }

        }

        public static void GetPrimaryIdFieldName(string _entityName, IOrganizationService service, out string _primaryField)
        {
            //Create RetrieveEntityRequest

            _primaryField = "";
            try
            {
                RetrieveEntityRequest retrievesEntityRequest = new RetrieveEntityRequest
                {
                    EntityFilters = EntityFilters.Entity,
                    LogicalName = _entityName
                };
                //Execute Request
                RetrieveEntityResponse retrieveEntityResponse = (RetrieveEntityResponse)service.Execute(retrievesEntityRequest);
                //_primaryKey = retrieveEntityResponse.EntityMetadata.PrimaryIdAttribute;
                _primaryField = retrieveEntityResponse.EntityMetadata.PrimaryNameAttribute;
            }
            catch (Exception ex)
            {
                _primaryField = "ERROR";
                _primaryField = ex.Message;
            }
        }
        public static void ActionOnApprovalofOCLPrice(IOrganizationService service, Entity ApprovalLine, ITracingService tracingService)
        {
            bool current_PendingForAction = true;
            bool isAnyPendingForActon = false;
            int approvalStatus = ApprovalLine.GetAttributeValue<OptionSetValue>("hil_approvalstatus").Value;
            EntityReference approvalHeader = ApprovalLine.GetAttributeValue<EntityReference>("hil_approval");
            if (approvalStatus == (int)ApprovalStatus.Approved || approvalStatus == (int)ApprovalStatus.Rejected)
            {
                QueryExpression _query = new QueryExpression("hil_oclproductapproval");
                _query.ColumnSet = new ColumnSet("hil_approvalstatus", "hil_oclproduct");
                _query.Criteria = new FilterExpression(LogicalOperator.And);
                _query.Criteria.AddCondition(new ConditionExpression("hil_approval", ConditionOperator.Equal, approvalHeader.Id));
                EntityCollection lineColl = service.RetrieveMultiple(_query);
                foreach (Entity entity in lineColl.Entities)
                {
                    if (entity.Contains("hil_approvalstatus"))
                    {
                        if (entity.GetAttributeValue<OptionSetValue>("hil_approvalstatus").Value != (int)ApprovalStatus.Approved && entity.GetAttributeValue<OptionSetValue>("hil_approvalstatus").Value != (int)ApprovalStatus.Rejected)
                            current_PendingForAction = false;
                    }
                    else
                        current_PendingForAction = false;
                    Guid oclPrdID = entity.GetAttributeValue<EntityReference>("hil_oclproduct").Id;
                    _query = new QueryExpression("hil_oclproductapproval");
                    _query.ColumnSet = new ColumnSet("hil_approvalstatus");
                    _query.Criteria = new FilterExpression(LogicalOperator.And);

                    FilterExpression filter1 = new FilterExpression(LogicalOperator.And);
                    _query.Criteria.AddCondition(new ConditionExpression("hil_oclproduct", ConditionOperator.Equal, oclPrdID));

                    FilterExpression filter2 = new FilterExpression(LogicalOperator.Or);
                    filter2.AddCondition("hil_approvalstatus", ConditionOperator.Equal, (int)ApprovalStatus.Submitted_for_Approval);
                    filter2.AddCondition("hil_approvalstatus", ConditionOperator.Equal, (int)ApprovalStatus.Send_for_Revision);

                    _query.Criteria.AddFilter(filter1);
                    _query.Criteria.AddFilter(filter2);

                    EntityCollection lineCollNew = service.RetrieveMultiple(_query);
                    if (lineCollNew.Entities.Count > 0)
                    {
                        isAnyPendingForActon = true;
                    }

                }

                if (current_PendingForAction && !isAnyPendingForActon)
                {
                    Entity approverHeaderEnt = service.Retrieve(approvalHeader.LogicalName, approvalHeader.Id, new ColumnSet("regardingobjectid"));
                    EntityReference targetEntityRef = approverHeaderEnt.GetAttributeValue<EntityReference>("regardingobjectid");
                    _query = new QueryExpression(approverHeaderEnt.LogicalName);
                    _query.ColumnSet = new ColumnSet("ownerid");
                    _query.Criteria = new FilterExpression(LogicalOperator.And);
                    _query.Criteria.AddCondition(new ConditionExpression("regardingobjectid", ConditionOperator.Equal, targetEntityRef.Id));
                    EntityCollection headerColl = service.RetrieveMultiple(_query);
                    List<Guid> approverList = new List<Guid>();
                    foreach (Entity approver in headerColl.Entities)
                    {
                        EntityReference owner = approver.GetAttributeValue<EntityReference>("ownerid");
                        approverList.Add(owner.Id);
                    }
                    submitForApproval(service, targetEntityRef, approverList,tracingService);
                }

            }
        }   
    }
    public class ApprovalsClass
    {
        public Entity ApprovalMatrixLineEntity { get; set; }
        public Guid ApproverID { get; set; }
        public string ApprovalName { get; set; }
        public string ApprovalEntity { get; set; }
        public int Level { get; set; }
        public Guid MAterialGroupID { get; set; }
        public string MAterialGroupName { get; set; }
        public string MAterialGroupEntity { get; set; }
    }
}
