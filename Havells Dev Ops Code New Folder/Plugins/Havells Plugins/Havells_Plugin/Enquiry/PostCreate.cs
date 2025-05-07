using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Havells_Plugin.Enquiry
{
    public class PostCreate:IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            string _selectedDivision = string.Empty;
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == "lead" && context.MessageName.ToUpper() == "CREATE")
                {
                    tracingService.Trace("1");
                    Entity entity = (Entity)context.InputParameters["Target"];
                    tracingService.Trace("12");
                    if (entity.Contains("leadsourcecode") && entity.GetAttributeValue<OptionSetValue>("leadsourcecode").Value == 11)
                    {
                        tracingService.Trace("2");

                        if (entity.Attributes.Contains("hil_selecteddivisions"))
                        {
                            _selectedDivision = entity.GetAttributeValue<string>("hil_selecteddivisions");
                            Entity obj_leadProduct = new Entity("hil_leadproduct");
                            string[] divisions;

                            if (_selectedDivision.IndexOf(',') < 0)
                            {
                                _selectedDivision = _selectedDivision + ",";
                            }

                            divisions = _selectedDivision.Split(',');

                            for (int i = 0; i < divisions.Length; i++)
                            {
                                if (divisions[i] != null && divisions[i].ToString().Trim().Length > 0)
                                {
                                    obj_leadProduct.Attributes["hil_enquiry"] = new EntityReference("lead", entity.Id);
                                    obj_leadProduct.Attributes["hil_productdivision"] = new EntityReference("product", Guid.Parse(divisions[i]));
                                    service.Create(obj_leadProduct);
                                }
                            }
                            Entity obj_lead1 = (Entity)service.Retrieve("lead", entity.Id, new ColumnSet("hil_leadstatus"));
                            obj_lead1.Attributes["hil_leadstatus"] = new OptionSetValue(2);
                            service.Update(obj_lead1);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                throw new InvalidPluginExecutionException("Error in Havells_Plugin.Enquiry.PostCreate " + e.Message);
            }
        }
    }
}
