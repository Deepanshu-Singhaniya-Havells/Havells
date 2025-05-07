using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Configuration;

namespace ClaimGeneration
{
    public class CommonLib
    {
        public string status { get; set; }
        public string statusRemarks { get; set; }

        public CommonLib CreateSAWActivity(Guid _jobId, decimal _amount, string _sawCategory, IOrganizationService _service, string _remarks, EntityReference _repeatRefjobId)
        {
            try
            {
                QueryExpression qryExp = new QueryExpression(msdyn_workorder.EntityLogicalName);
                qryExp.ColumnSet = new ColumnSet("hil_owneraccount", "ownerid");
                qryExp.Criteria = new FilterExpression(LogicalOperator.And);
                qryExp.Criteria.AddCondition("msdyn_workorderid", ConditionOperator.Equal, _jobId);
                EntityCollection entCol = _service.RetrieveMultiple(qryExp);
                if (entCol.Entities.Count > 0)
                {
                    if (!CheckIfSAWActivityExist(_jobId, _service, _sawCategory))
                    {
                        Entity entSAWActivity = new Entity("hil_sawactivity");
                        entSAWActivity["hil_sawcategory"] = new EntityReference("hil_serviceactionwork", new Guid(_sawCategory));
                        entSAWActivity["hil_relatedchannelpartner"] = entCol.Entities[0].GetAttributeValue<EntityReference>("hil_owneraccount");
                        entSAWActivity["hil_jobid"] = new EntityReference("msdyn_workorder", _jobId);
                        entSAWActivity["hil_description"] = _remarks;
                        if (_repeatRefjobId != null)
                        {
                            entSAWActivity["hil_repeatreferencejob"] = _repeatRefjobId;
                        }
                        entSAWActivity["hil_amount"] = new Money(_amount);
                        entSAWActivity["hil_approvalstatus"] = new OptionSetValue(1); //requested
                        Guid sawActivityId = _service.Create(entSAWActivity);
                        CreateSAWActivityApprovals(sawActivityId, _service);
                    }
                    return new CommonLib() { status = "200", statusRemarks = "OK" };
                }
                else
                {
                    return new CommonLib() { status = "204", statusRemarks = "something went wrong." };
                }
            }
            catch (Exception ex)
            {
                return new CommonLib() { status = "204", statusRemarks = ex.Message };
            }
        }

        public CommonLib CreateSAWActivityApprovals(Guid _sawActivityId, IOrganizationService _service)
        {
            try
            {
                Entity ent = _service.Retrieve("hil_sawactivity", _sawActivityId, new ColumnSet("hil_sawcategory", "hil_jobid"));
                if (ent != null)
                {
                    EntityReference _salesOffice = null;
                    EntityReference _productCatg = null;
                    EntityReference _job = null;
                    EntityReference _picUser = null;
                    EntityReference _picPosition = null;
                    Entity entJob = _service.Retrieve("msdyn_workorder", ent.GetAttributeValue<EntityReference>("hil_jobid").Id, new ColumnSet("hil_salesoffice", "hil_productcategory"));
                    if (entJob != null)
                    {
                        if (entJob.Attributes.Contains("hil_salesoffice"))
                        {
                            _salesOffice = entJob.GetAttributeValue<EntityReference>("hil_salesoffice");
                            _productCatg = entJob.GetAttributeValue<EntityReference>("hil_productcategory");
                            _job = ent.GetAttributeValue<EntityReference>("hil_jobid");
                        }
                    }
                    if (_salesOffice != null && _productCatg != null)
                    {
                        QueryExpression qrySBU = new QueryExpression("hil_sbubranchmapping");
                        qrySBU.ColumnSet = new ColumnSet("hil_nsh", "hil_nph", "hil_branchheaduser");
                        qrySBU.Criteria = new FilterExpression(LogicalOperator.And);
                        qrySBU.Criteria.AddCondition("hil_salesoffice", ConditionOperator.Equal, _salesOffice.Id);
                        qrySBU.Criteria.AddCondition("hil_productdivision", ConditionOperator.Equal, _productCatg.Id);
                        EntityCollection entColSBU = _service.RetrieveMultiple(qrySBU);

                        QueryExpression qryExp = new QueryExpression("hil_sawcategoryapprovals");
                        qryExp.ColumnSet = new ColumnSet("hil_picuser", "hil_picposition", "hil_level");
                        qryExp.Criteria = new FilterExpression(LogicalOperator.And);
                        qryExp.Criteria.AddCondition("hil_sawcategoryid", ConditionOperator.Equal, ent.GetAttributeValue<EntityReference>("hil_sawcategory").Id);
                        EntityCollection entCol = _service.RetrieveMultiple(qryExp);
                        if (entCol.Entities.Count > 0)
                        {
                            foreach (Entity entTemp in entCol.Entities)
                            {
                                if (entTemp.Attributes.Contains("hil_picuser"))
                                {
                                    _picUser = entTemp.GetAttributeValue<EntityReference>("hil_picuser");
                                }
                                else
                                {
                                    _picPosition = entTemp.GetAttributeValue<EntityReference>("hil_picposition");
                                    if (_picPosition.Name.ToUpper() == "BSH")
                                    {
                                        if (entColSBU.Entities.Count > 0)
                                        {
                                            _picUser = entColSBU.Entities[0].GetAttributeValue<EntityReference>("hil_branchheaduser");
                                        }
                                    }
                                    else if (_picPosition.Name.ToUpper() == "NSH")
                                    {
                                        _picUser = entColSBU.Entities[0].GetAttributeValue<EntityReference>("hil_nsh");

                                    }
                                    else if (_picPosition.Name.ToUpper() == "NPH")
                                    {
                                        _picUser = entColSBU.Entities[0].GetAttributeValue<EntityReference>("hil_nph");
                                    }
                                }
                                Entity entSAWActivity = new Entity("hil_sawactivityapproval");
                                entSAWActivity["hil_sawactivity"] = new EntityReference("hil_sawactivity", _sawActivityId);
                                entSAWActivity["hil_jobid"] = _job;
                                entSAWActivity["hil_level"] = entTemp.GetAttributeValue<OptionSetValue>("hil_level");
                                if (entTemp.GetAttributeValue<OptionSetValue>("hil_level").Value == 1)
                                {
                                    entSAWActivity["hil_isenabled"] = true;
                                }
                                entSAWActivity["hil_approver"] = _picUser;
                                entSAWActivity["hil_approvalstatus"] = new OptionSetValue(1); //requested
                                Guid sawActivityId = _service.Create(entSAWActivity);
                            }
                            return new CommonLib() { status = "200", statusRemarks = "OK" };
                        }
                        else
                        {
                            return new CommonLib() { status = "204", statusRemarks = "SAW Activity approvals not found." };
                        }
                    }
                    else
                    {
                        return new CommonLib() { status = "204", statusRemarks = "Sales Office not found in Job." };
                    }
                }
                else
                {
                    return new CommonLib() { status = "204", statusRemarks = "Something went wrong." };
                }
            }
            catch (Exception ex)
            {
                return new CommonLib() { status = "204", statusRemarks = ex.Message };
            }
        }

        public bool CheckIfSAWActivityExist(Guid _jobId, IOrganizationService _service, string _sawCategory)
        {
            try
            {
                QueryExpression qryExp = new QueryExpression("hil_sawactivity");
                qryExp.ColumnSet = new ColumnSet("hil_jobid");
                qryExp.Criteria = new FilterExpression(LogicalOperator.And);
                qryExp.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, _jobId);
                qryExp.Criteria.AddCondition("hil_sawcategory", ConditionOperator.Equal, new Guid(_sawCategory));
                EntityCollection entCol = _service.RetrieveMultiple(qryExp);
                if (entCol.Entities.Count > 0) { return true; } else { return false; }
            }
            catch
            {
                return false;
            }
        }
    }

    public class SAWCategoryConst
    {
        public const string _GasChargePostAudit = "f5784d62-4bdd-ea11-a813-000d3af0563c";
        public const string _GasChargePreAuditPastInstallationHistory = "9d01bc73-c5db-ea11-a813-000d3af055b6";
        public const string _GasChargePreAuditPastRepeatHistory = "a8e50e38-4bdd-ea11-a813-000d3af0563c";
        public const string _KKGFailureReview = "e123bd08-8add-ea11-a813-000d3af05a4b";
        public const string _LocalPurchase = "d0a3babb-f3d0-ea11-a813-000d3af05a4b";
        public const string _LocalRepair = "d577a3c8-f3d0-ea11-a813-000d3af05a4b";
        public const string _ProductTransportation = "db074fae-f3d0-ea11-a813-000d3af05a4b";
        public const string _RepeatRepair = "b0918d74-44ed-ea11-a815-000d3af05d7b";
        public const string _OneTimeLaborException = "ad96a922-0aee-ea11-a815-000d3af057dd";
        public const string _AMCSpecialDiscount = "8ce750f5-d539-eb11-a813-0022486eaccc";
    }
}
