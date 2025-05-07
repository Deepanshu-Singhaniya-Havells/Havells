using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Havells_Plugin.Campaign
{
    public class PostUpdate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            try
            {
                if (context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == "campaign"
                    && context.MessageName.ToUpper() == "UPDATE" && context.Depth<2)
                {
                    Entity entityCampaign  = (Entity)context.InputParameters["Target"];

                    QueryExpression qrExp = new QueryExpression("campaign");
                    qrExp.ColumnSet = new ColumnSet(true);
                    ConditionExpression condExp = new ConditionExpression("campaignid", ConditionOperator.Equal, entityCampaign.Id);
                    qrExp.Criteria.AddCondition(condExp);
                    qrExp.NoLock = true;
                    EntityCollection collect_user = service.RetrieveMultiple(qrExp);

                    Entity entity = collect_user.Entities[0];

                    OptionSetValue campaignMedium = entity.GetAttributeValue<OptionSetValue>("hil_campaignmedium");
                    OptionSetValue hil_campaignstatus = entity.GetAttributeValue<OptionSetValue>("hil_campaignstatus");
                    string selecteddivisions = string.Empty;
                    string selectedenquirytypes = string.Empty;
                    tracingService.Trace("1");
                    if (hil_campaignstatus.Value == 1) // Campaign Approved
                    {
                        selecteddivisions = getCampaignProductDivision(service, entity);
                        tracingService.Trace(selecteddivisions);
                        selectedenquirytypes = getCampaignEnquiryType(service, entity);
                        tracingService.Trace(selectedenquirytypes);

                        string campaignCode = entity.GetAttributeValue<string>("codename");
                        string campaignBaseURL = entity.GetAttributeValue<string>("hil_campaignbaseurl");
                        bool selectesalldivisions = entity.GetAttributeValue<bool>("hil_selectesalldivisions");
                        bool selectesallenquirytypes = entity.GetAttributeValue<bool>("hil_selectesallenquirytypes");
                        bool displaydivisioninwebsite = entity.GetAttributeValue<bool>("hil_displaydivisioninwebsite");
                        bool displayenquirytypeinwebsite = entity.GetAttributeValue<bool>("hil_displayenquirytypeinwebsite");

                        if (selecteddivisions == string.Empty && !selectesalldivisions) {
                            throw new InvalidPluginExecutionException(" ***Campaign Product Division is required*** ");
                        }
                        else if (selectedenquirytypes == string.Empty && !selectesallenquirytypes)
                        {
                            throw new InvalidPluginExecutionException(" ***Campaign Enquiry Type is required*** ");
                        }
                        tracingService.Trace(campaignBaseURL);
                        if (campaignBaseURL != null)// Generate Campaign URL is Campaign Medium is not equal to Email/SMS
                        {
                            campaignBaseURL = campaignBaseURL + "?utm_campaigncode=" + campaignCode;
                            if (selectesalldivisions)
                            {
                                campaignBaseURL = campaignBaseURL + "&utm_campaigncontent=All";
                            }
                            else
                            {
                                campaignBaseURL = campaignBaseURL + "&utm_campaigncontent=" + selecteddivisions;
                            }
                            if (selectesallenquirytypes)
                            {
                                campaignBaseURL = campaignBaseURL + "&utm_campaigntype=All";
                            }
                            else
                            {
                                campaignBaseURL = campaignBaseURL + "&utm_campaigntype=" + selectedenquirytypes;
                            }
                            if (displaydivisioninwebsite)
                            {
                                campaignBaseURL = campaignBaseURL + "&utm_displaydivision=Yes";
                            }
                            else
                            {
                                campaignBaseURL = campaignBaseURL + "&utm_displaydivision=No";
                            }
                            if (displayenquirytypeinwebsite)
                            {
                                campaignBaseURL = campaignBaseURL + "&utm_displayenquirytype=Yes";
                            }
                            else
                            {
                                campaignBaseURL = campaignBaseURL + "&utm_displayenquirytype=No";
                            }
                            Entity entCampaign = new Entity("campaign");
                            entCampaign.Id = entity.Id;
                            entCampaign["hil_selecteddivisions"] = selecteddivisions;
                            entCampaign["hil_selectedenquirytypes"] = selectedenquirytypes;
                            entCampaign["hil_campaignurl"] = campaignBaseURL;
                            service.Update(entCampaign);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.Campaign.PostUpdate.Execute  ***" + ex.Message + "***  ");
            }
        }

        public string getCampaignProductDivision(IOrganizationService service, Entity ent)
        {
            string _retValue = string.Empty;
            try
            {
                LinkEntity lnkEntDivision = new LinkEntity
                {
                    LinkFromEntityName = "hil_campaigndivisions",
                    LinkToEntityName = "product",
                    LinkFromAttributeName = "hil_divisionid",
                    LinkToAttributeName = "productid",
                    Columns = new ColumnSet("hil_sapcode"),
                    EntityAlias = "prod",
                    JoinOperator = JoinOperator.Inner
                };

                QueryExpression qrExp = new QueryExpression("hil_campaigndivisions");
                qrExp.ColumnSet = new ColumnSet("hil_divisionid");
                ConditionExpression condExp = new ConditionExpression("hil_campaignid", ConditionOperator.Equal, ent.Id);
                qrExp.Criteria.AddCondition(condExp);
                qrExp.NoLock = true;
                qrExp.LinkEntities.Add(lnkEntDivision);
                EntityCollection collect_user = service.RetrieveMultiple(qrExp);
                if (collect_user.Entities.Count > 0)
                {
                    int i = 1;
                    foreach (Entity entDiv in collect_user.Entities)
                    {
                        if (i != collect_user.Entities.Count)
                        {
                            _retValue = _retValue + entDiv.GetAttributeValue<AliasedValue>("prod.hil_sapcode").Value.ToString() + ",";
                        }
                        else
                        {
                            _retValue = _retValue + entDiv.GetAttributeValue<AliasedValue>("prod.hil_sapcode").Value.ToString();
                        }
                        i += 1;
                    }
                }
                return _retValue;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.Campaign.PostUpdate.getCampaignProductDivision  ***" + ex.Message + "***  ");
            }
        }

        public string getCampaignEnquiryType(IOrganizationService service, Entity ent)
        {
            string _retValue = string.Empty;
            try
            {
                LinkEntity lnkEntEnquiryType = new LinkEntity
                {
                    LinkFromEntityName = "hil_campaignenquirytypes",
                    LinkToEntityName = "hil_enquirytype",
                    LinkFromAttributeName = "hil_enquirytypeid",
                    LinkToAttributeName = "hil_enquirytypeid",
                    Columns = new ColumnSet("hil_enquirytypecode"),
                    EntityAlias = "enq",
                    JoinOperator = JoinOperator.Inner
                };

                QueryExpression qrExp = new QueryExpression("hil_campaignenquirytypes");
                qrExp.ColumnSet = new ColumnSet("hil_enquirytypeid");
                ConditionExpression condExp = new ConditionExpression("hil_campaignid", ConditionOperator.Equal, ent.Id);
                qrExp.Criteria.AddCondition(condExp);
                qrExp.NoLock = true;
                qrExp.LinkEntities.Add(lnkEntEnquiryType);
                EntityCollection collect_user = service.RetrieveMultiple(qrExp);
                if (collect_user.Entities.Count > 0)
                {
                    int i = 1;
                    foreach (Entity entDiv in collect_user.Entities)
                    {
                        if (i != collect_user.Entities.Count)
                        {
                            _retValue = _retValue + entDiv.GetAttributeValue<AliasedValue>("enq.hil_enquirytypecode").Value.ToString() + ",";
                        }
                        else
                        {
                            _retValue = _retValue + entDiv.GetAttributeValue<AliasedValue>("enq.hil_enquirytypecode").Value.ToString();
                        }
                        i += 1;
                    }
                }
                return _retValue;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.Campaign.PostUpdate.getCampaignEnquiryType  ***" + ex.Message + "***  ");
            }
        }
    }
}
