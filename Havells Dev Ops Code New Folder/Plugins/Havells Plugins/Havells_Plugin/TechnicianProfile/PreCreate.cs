using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Havells_Plugin.TechnicianProfile
{
    public class PreCreate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region Execute Main Region
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService traceservice = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            Entity entity = (Entity)context.InputParameters["Target"];
            #endregion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                && context.PrimaryEntityName.ToLower() == "hil_technician" && context.MessageName.ToUpper() == "CREATE")
                {
                    EntityReference ownerId = entity.GetAttributeValue<EntityReference>("ownerid");
                    EntityReference partnerdepartmentId = entity.GetAttributeValue<EntityReference>("hil_partnerdepartment");
                    string firstName = entity.GetAttributeValue<string>("hil_firstname");
                    string lastName = entity.GetAttributeValue<string>("hil_lastname");

                    QueryExpression qsUser = new QueryExpression("systemuser");
                    qsUser.ColumnSet = new ColumnSet("systemuserid", "positionid", "parentsystemuserid");
                    ConditionExpression condExp = new ConditionExpression("systemuserid", ConditionOperator.Equal, ownerId.Id);
                    qsUser.Criteria.AddCondition(condExp);

                    EntityCollection collect_user = service.RetrieveMultiple(qsUser);
                    if (collect_user.Entities.Count == 1)
                    {
                        if (!collect_user.Entities[0].Contains("positionid"))
                        {
                            throw new InvalidPluginExecutionException("Havells_Plugin.FSEProfile.PreCreate.Execute  *** Position is not set !!! Please contact to System Admin.***  ");
                        }
                        else if (!collect_user.Entities[0].Contains("parentsystemuserid"))
                        {
                            throw new InvalidPluginExecutionException("Havells_Plugin.FSEProfile.PreCreate.Execute  *** Manager is not set !!! Please contact to System Admin.***  ");
                        }
                        else if (collect_user.Entities[0].GetAttributeValue<EntityReference>("positionid").Name.ToUpper() != "FRANCHISE" && collect_user.Entities[0].GetAttributeValue<EntityReference>("positionid").Name.ToUpper() != "DSE")
                        {
                            throw new InvalidPluginExecutionException("Havells_Plugin.FSEProfile.PreCreate.Execute  *** Access denied !!! Please contact to System Admin.***  ");
                        }
                        EntityReference er = GetManagerId(service, collect_user.Entities[0].Id);
                        entity["hil_approver"] = er;
                        entity["hil_manageremployeeid"] = GetApproverEmployeeID(service, er.Id);
                        entity["hil_technicianuseridrequest"] = GetADIdRequestOwner(service, partnerdepartmentId.Id);
                        entity["hil_salesoffice"] = GetSalesOffice(service, er.Id);
                        entity["hil_aspcode"] = GetASPCode(service, ownerId.Id);
                        entity["hil_firstname"] = firstName.ToUpper();
                        if (lastName != null && lastName != string.Empty)
                        {
                            entity["hil_lastname"] = lastName.ToUpper();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.FSEProfile.PreCreate.Execute  ***" + ex.Message + "***  ");
            }
        }

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
                        throw new InvalidPluginExecutionException("Havells_Plugin.FSEProfile.PreCreate.GetManagerId  *** Manager's Position is not defined !!! Please contact to System Admin.***  ");
                    }
                    else if (ecObj.Entities[0].GetAttributeValue<EntityReference>("positionid").Name.ToUpper() != "BSH")
                    {
                        if (!ecObj.Entities[0].Contains("parentsystemuserid"))
                        {
                            throw new InvalidPluginExecutionException("Havells_Plugin.FSEProfile.PreCreate.GetManagerId  *** Manager's Hierarchy is not defined !!! Please contact to System Admin.***  ");
                        }
                        else
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
                throw new InvalidPluginExecutionException("Havells_Plugin.WorkOrder.PostUpdate.Execute.GetManagerId: " + ex.Message);
            }
            return _managerRecord;
        }

        public static EntityReference GetADIdRequestOwner(IOrganizationService service, Guid partnerDepartmentGuId)
        {
            EntityReference _adRequestOwner = null;
            try
            {
                Entity ent = service.Retrieve("hil_partnerdepartment", partnerDepartmentGuId, new ColumnSet("hil_adprofilercreator"));
                if (ent != null)
                {
                    if (ent.Contains("hil_adprofilercreator"))
                    {
                        return ent.GetAttributeValue<EntityReference>("hil_adprofilercreator");
                    }
                    else
                    {
                        throw new InvalidPluginExecutionException("Havells_Plugin.FSEProfile.PreCreate.GetADIdRequestOwner  *** AD ID Request Owner is not defined !!! Please contact to System Admin.***  ");
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.WorkOrder.PostUpdate.Execute.GetADRequestOwner: " + ex.Message);
            }
            return _adRequestOwner;
        }

        public static EntityReference GetSalesOffice(IOrganizationService service, Guid managerId)
        {
            EntityReference _branch = null;
            try
            {
                QueryExpression qsBranch = new QueryExpression("hil_sbubranchmapping");
                qsBranch.ColumnSet = new ColumnSet("hil_salesoffice");
                ConditionExpression condIngeConfg = new ConditionExpression("hil_branchheaduser", ConditionOperator.Equal, managerId);
                qsBranch.Criteria.AddCondition(condIngeConfg);
                condIngeConfg = new ConditionExpression("statecode", ConditionOperator.Equal, 0);
                qsBranch.Criteria.AddCondition(condIngeConfg);
                qsBranch.Orders.Add(new OrderExpression("createdon", OrderType.Ascending));

                EntityCollection ecObj = service.RetrieveMultiple(qsBranch);
                if (ecObj.Entities.Count >= 1)
                {
                    if (ecObj.Entities[0].Contains("hil_salesoffice"))
                    {
                        _branch = ecObj.Entities[0].GetAttributeValue<EntityReference>("hil_salesoffice");
                    }
                    else
                    {
                        throw new InvalidPluginExecutionException("Havells_Plugin.FSEProfile.PreCreate.GetSalesOffice  *** Approver Branch is not defined !!! Please contact to System Admin.***  ");
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.WorkOrder.PostUpdate.Execute.GetSalesOffice: " + ex.Message);
            }
            return _branch;
        }

        public static string GetASPCode(IOrganizationService service, Guid ownerId)
        {
            string _aspCode = string.Empty;
            try
            {
                string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>" +
                "<entity name='systemuser'><attribute name='fullname' /><filter type='and'>" +
                "<condition attribute='systemuserid' operator='eq' value='{" + ownerId.ToString() + @"}' />" +
                "</filter>" +
                "<link-entity name='account' from='accountid' to='hil_account' link-type='inner' alias='aa'>" +
                "<attribute name='accountnumber' />" +
                "</link-entity>" +
                "</entity>" +
                "</fetch>";
                EntityCollection ecObj = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (ecObj.Entities.Count == 1)
                {
                    if (ecObj.Entities[0].Contains("aa.accountnumber"))
                    {
                        _aspCode = ecObj.Entities[0].GetAttributeValue<AliasedValue>("aa.accountnumber").Value.ToString();
                    }
                    else
                    {
                        throw new InvalidPluginExecutionException("Havells_Plugin.FSEProfile.PreCreate.GetASPCode  *** ASP Code is not defined !!! Please contact to System Admin.***  ");
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.WorkOrder.PostUpdate.Execute.GetASPCode: " + ex.Message);
            }
            return _aspCode;
        }

        public static string GetApproverEmployeeID(IOrganizationService service, Guid managerId)
        {
            string _employeeId = string.Empty;
            try
            {
                string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>" +
                "<entity name='systemuser'><attribute name='hil_employeecode' /><filter type='and'>" +
                "<condition attribute='systemuserid' operator='eq' value='{" + managerId.ToString() + @"}' />" +
                "</filter>" +
                "</entity>" +
                "</fetch>";
                EntityCollection ecObj = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (ecObj.Entities.Count == 1)
                {
                    if (ecObj.Entities[0].Contains("hil_employeecode"))
                    {
                        _employeeId = ecObj.Entities[0].GetAttributeValue<string>("hil_employeecode");
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.WorkOrder.PostUpdate.Execute.GetApproverEmployeeID: " + ex.Message);
            }
            return _employeeId;
        }
    }
}
