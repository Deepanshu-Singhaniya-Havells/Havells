using System;
using System.Collections.Generic;
using HavellsNewPlugin.CustomerAsset;
using HavellsNewPlugin.Helper;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.WorkOrder
{
    public class WorkOderPostUpdate_Validation : IPlugin
    {
        public static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            #region MainRegion
            try
            {
                if (context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is Entity
                    && context.MessageName.ToUpper() == "UPDATE")
                {
                    tracingService.Trace("1");
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity entity = (Entity)context.InputParameters["Target"];
                    tracingService.Trace("1");
                    bool throwError = false;
                    Guid user = context.UserId;
                    Entity preImageCA = ((Entity)context.PreEntityImages["image"]);
                    tracingService.Trace("2");
                    if (entity.Contains("hil_customerref") && preImageCA.Contains("hil_customerref"))
                    {
                        tracingService.Trace("3");
                        EntityReference customer = entity.GetAttributeValue<EntityReference>("hil_customerref");
                        EntityReference customerImg = preImageCA.GetAttributeValue<EntityReference>("hil_customerref");
                        if (customer.Id != customerImg.Id)
                        {
                            tracingService.Trace("4");
                            throwError = true;
                        }
                        tracingService.Trace("5");
                    }
                    tracingService.Trace("6");
                    //if (entity.Contains("hil_address"))
                    //{
                    //    EntityReference Address = entity.GetAttributeValue<EntityReference>("hil_address");
                    //    EntityReference AddressImg = preImageCA.GetAttributeValue<EntityReference>("hil_address");
                    //    if (Address.Id != AddressImg.Id)
                    //    {
                    //        throwError = true;
                    //    }
                    //}
                    if (entity.Contains("hil_consumertype") && preImageCA.Contains("hil_customerref"))
                    {
                        tracingService.Trace("7");
                        EntityReference customerType = entity.GetAttributeValue<EntityReference>("hil_consumertype");
                        tracingService.Trace("7.1");
                        if (preImageCA.Contains("hil_customerref"))
                        {
                            tracingService.Trace("7.2");
                            EntityReference customerTypeImg = preImageCA.GetAttributeValue<EntityReference>("hil_consumertype");
                            if (customerTypeImg != null)
                            {
                                tracingService.Trace("7.3" + customerTypeImg);
                                tracingService.Trace("7.4" + customerTypeImg.Name);
                                tracingService.Trace("7.5" + customerTypeImg.Id);
                                if (customerType.Id != customerTypeImg.Id)
                                {
                                    throwError = true;
                                    tracingService.Trace("8");
                                }
                            }

                        }
                        tracingService.Trace("9");
                    }
                    tracingService.Trace("10");
                    if (entity.Contains("hil_consumercategory") && preImageCA.Contains("hil_consumercategory"))
                    {
                        tracingService.Trace("11");
                        EntityReference customerCategory = entity.GetAttributeValue<EntityReference>("hil_consumercategory");
                        EntityReference customerCategoryImg = preImageCA.GetAttributeValue<EntityReference>("hil_consumercategory");
                        if (customerCategory.Id != customerCategoryImg.Id)
                        {
                            tracingService.Trace("12");
                            throwError = true;
                        }
                        tracingService.Trace("13");
                    }
                    tracingService.Trace("14");
                    if (entity.Contains("hil_callertype") && preImageCA.Contains("hil_callertype"))
                    {
                        tracingService.Trace("15");
                        OptionSetValue CallerType = entity.GetAttributeValue<OptionSetValue>("hil_callertype");
                        OptionSetValue callerTypeImg = preImageCA.GetAttributeValue<OptionSetValue>("hil_callertype");
                        if (CallerType.Value != callerTypeImg.Value)
                        {
                            tracingService.Trace("16");
                            throwError = true;
                        }
                        tracingService.Trace("17");
                    }
                    tracingService.Trace("18");
                    if (entity.Contains("hil_newserialnumber") && preImageCA.Contains("hil_newserialnumber"))
                    {
                        tracingService.Trace("19");
                        String _emp_ASP_D = entity.GetAttributeValue<String>("hil_newserialnumber");
                        String _emp_ASP_DImg = preImageCA.GetAttributeValue<String>("hil_newserialnumber");
                        if (_emp_ASP_D != _emp_ASP_DImg)
                        {
                            tracingService.Trace("20");
                            throwError = true;
                        }
                        tracingService.Trace("21");
                    }
                    tracingService.Trace("22");
                    if (entity.Contains("hil_preferreddate") && preImageCA.Contains("hil_preferreddate"))
                    {
                        tracingService.Trace("23");
                        DateTime preferDate = entity.GetAttributeValue<DateTime>("hil_preferreddate");
                        DateTime preferDateImg = preImageCA.GetAttributeValue<DateTime>("hil_preferreddate");
                        if (preferDate != preferDateImg)
                        {
                            tracingService.Trace("24");
                            throwError = true;
                        }
                        tracingService.Trace("25");
                    }
                    tracingService.Trace("26");
                    if (entity.Contains("hil_preferredtime") && preImageCA.Contains("hil_preferredtime"))
                    {
                        tracingService.Trace("27");
                        OptionSetValue preferdTime = entity.GetAttributeValue<OptionSetValue>("hil_preferredtime");
                        OptionSetValue preferdTimeImg = preImageCA.GetAttributeValue<OptionSetValue>("hil_preferredtime");
                        if (preferdTime.Value != preferdTimeImg.Value)
                        {
                            tracingService.Trace("28");
                            throwError = true;
                        }
                        tracingService.Trace("29");
                    }
                    tracingService.Trace("30");
                    if (throwError)
                    {
                        tracingService.Trace("31");
                        string position = HelperClass.getUserPosition(user, service, tracingService);
                        tracingService.Trace("31");
                        if (position == "DSE" || position == "Franchise" || position == "Franchise Technician" || position == null)
                            throw new InvalidPluginExecutionException("You are not authorized to update this detail, please contact to your Branch Office.");
                        tracingService.Trace("32");
                    }
                    tracingService.Trace("33");
                    if (entity.Contains("hil_mobilenumber") && preImageCA.Contains("hil_mobilenumber"))
                    {
                        tracingService.Trace("34");
                        String hil_mobilenumber = entity.GetAttributeValue<String>("hil_mobilenumber");
                        String hil_mobilenumberImg = preImageCA.GetAttributeValue<String>("hil_mobilenumber");
                        tracingService.Trace("35");
                        if (hil_mobilenumber != hil_mobilenumberImg)
                        {
                            tracingService.Trace("36");
                            throw new InvalidPluginExecutionException("You are not authorized to update this detail.");
                        }
                        tracingService.Trace("37");
                    }
                    if (preImageCA.Contains("msdyn_customerasset"))
                    {
                        Entity customerAsset = service.Retrieve(preImageCA.GetAttributeValue<EntityReference>("msdyn_customerasset").LogicalName,
                            preImageCA.GetAttributeValue<EntityReference>("msdyn_customerasset").Id, new ColumnSet("hil_customer"));
                        EntityReference customerAssetCustomer = customerAsset.GetAttributeValue<EntityReference>("hil_customer");
                        EntityReference jobCustomer = preImageCA.GetAttributeValue<EntityReference>("hil_customerref");
                        if (jobCustomer.Id != customerAssetCustomer.Id)
                        {
                            throw new InvalidPluginExecutionException("Customer Asset doesn't belong to " + jobCustomer.Name);
                        }
                    }

                    tracingService.Trace("38");
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error: " + ex.Message);
            }
            #endregion
        }

    }
}
