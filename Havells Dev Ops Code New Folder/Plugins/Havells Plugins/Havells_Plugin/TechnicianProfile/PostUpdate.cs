using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Havells_Plugin.TechnicianProfile
{
    public class PostUpdate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region Execute Main Region
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService traceservice = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            Entity entity = (Entity)context.InputParameters["Target"];
            Entity entTechnicianInfo = null;
            Entity pPinCodeInfo = null;
            Entity cPinCodeInfo = null;

            #endregion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                && context.PrimaryEntityName.ToLower() == "hil_technician" && context.MessageName.ToUpper() == "UPDATE" && context.Depth < 2)
                {
                    OptionSetValue approvalStatus = entity.GetAttributeValue<OptionSetValue>("hil_technicianstatus");
                    if (approvalStatus.Value == 3) //Approved by BSH
                    {
                        entTechnicianInfo = service.Retrieve("hil_technician", entity.Id, new ColumnSet("hil_aspcode","hil_approver", "hil_manageremployeeid", "hil_sameaspermanentaddress", "hil_paddressline1", "hil_paddressline2", "hil_plandmark", "hil_ppincode", "hil_firstname", "hil_lastname", "hil_caddressline1", "hil_caddressline2", "hil_clandmark", "hil_cpincode"));

                        if (entTechnicianInfo.Attributes.Contains("hil_ppincode") || entTechnicianInfo.Contains("hil_ppincode"))
                        {
                            pPinCodeInfo = service.Retrieve("hil_businessmapping", entTechnicianInfo.GetAttributeValue<EntityReference>("hil_ppincode").Id, new ColumnSet("hil_city", "hil_pincode", "hil_state"));
                        }
                        if (entTechnicianInfo.Attributes.Contains("hil_cpincode") || entTechnicianInfo.Contains("hil_cpincode"))
                        {
                            cPinCodeInfo = service.Retrieve("hil_businessmapping", entTechnicianInfo.GetAttributeValue<EntityReference>("hil_cpincode").Id, new ColumnSet("hil_city", "hil_pincode", "hil_state"));
                        }

                        string managerEmployeeId = entTechnicianInfo.GetAttributeValue<string>("hil_manageremployeeid");
                        EntityReference manager = entTechnicianInfo.GetAttributeValue<EntityReference>("hil_approver");
                        if (managerEmployeeId != null && managerEmployeeId.Trim().Length > 0 && manager != null) //Submitted
                        {
                            SystemUser ent = new SystemUser();
                            ent.Id = manager.Id;
                            ent.hil_EmployeeCode = managerEmployeeId;
                            service.Update(ent);

                            if (entTechnicianInfo.Attributes.Contains("hil_userid") || entTechnicianInfo.Contains("hil_userid"))
                            {
                                ent.FirstName = entTechnicianInfo.GetAttributeValue<string>("hil_firstname");
                                ent.LastName = entTechnicianInfo.GetAttributeValue<string>("hil_lastname");

                                ent.Address1_Line1 = entTechnicianInfo.GetAttributeValue<string>("hil_paddressline1");
                                ent.Address1_Line2 = entTechnicianInfo.GetAttributeValue<string>("hil_paddressline2");
                                ent.Address1_Line3 = entTechnicianInfo.GetAttributeValue<string>("hil_plandmark");

                                if (pPinCodeInfo != null)
                                {
                                    ent.Address1_City = pPinCodeInfo.GetAttributeValue<EntityReference>("hil_city").Name;
                                    ent.Address1_PostalCode = pPinCodeInfo.GetAttributeValue<EntityReference>("hil_pincode").Name;
                                    ent.Address1_StateOrProvince = pPinCodeInfo.GetAttributeValue<EntityReference>("hil_state").Name;
                                    ent.Address1_Country = "INDIA";
                                }
                                ent.Address2_Line1 = entTechnicianInfo.GetAttributeValue<string>("hil_caddressline1");
                                ent.Address2_Line2 = entTechnicianInfo.GetAttributeValue<string>("hil_caddressline2");
                                ent.Address2_Line3 = entTechnicianInfo.GetAttributeValue<string>("hil_clandmark");

                                if (cPinCodeInfo != null)
                                {
                                    ent.Address2_City = cPinCodeInfo.GetAttributeValue<EntityReference>("hil_city").Name;
                                    ent.Address2_PostalCode = cPinCodeInfo.GetAttributeValue<EntityReference>("hil_pincode").Name;
                                    ent.Address2_StateOrProvince = cPinCodeInfo.GetAttributeValue<EntityReference>("hil_state").Name;
                                    ent.Address2_Country = "INDIA";
                                }
                                service.Update(ent);
                            }
                        }

                        string care360Id = GenerateCare360Id(service, entTechnicianInfo.GetAttributeValue<string>("hil_aspcode"));
                        if (care360Id != string.Empty)
                        {
                            Entity entTechnician = new Entity("hil_technician");
                            entTechnician.Id = entity.Id;
                            entTechnician["hil_care360id"] = care360Id;
                            service.Update(entTechnician);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.TechnicianProfile.PostUpdate.Execute  ***" + ex.Message + "***  ");
            }
        }

        private string GenerateCare360Id(IOrganizationService service,string aspCode) {
            string _retValue = string.Empty;
            int technicianCount = 0;
            try
            {
                QueryExpression qsUser = new QueryExpression("hil_technician");
                qsUser.ColumnSet = new ColumnSet("hil_aspcode");
                ConditionExpression condExp = new ConditionExpression("hil_aspcode", ConditionOperator.Equal, aspCode);
                qsUser.Criteria.AddCondition(condExp);
                condExp = new ConditionExpression("hil_care360id", ConditionOperator.NotNull);
                qsUser.Criteria.AddCondition(condExp);
                qsUser.NoLock = true;
                EntityCollection collect_user = service.RetrieveMultiple(qsUser);
                if (collect_user.Entities.Count > 0) {
                    technicianCount = collect_user.Entities.Count + 1;
                }
                _retValue = aspCode + technicianCount.ToString().PadLeft(3, '0');
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.TechnicianProfile.PostUpdate.GenerateCare360Id  ***" + ex.Message + "***  ");
            }
            return _retValue;
        }
    }
}
