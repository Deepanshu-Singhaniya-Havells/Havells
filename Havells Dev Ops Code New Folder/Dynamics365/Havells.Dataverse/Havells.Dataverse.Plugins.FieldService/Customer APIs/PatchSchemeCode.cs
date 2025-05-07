using Havells.Dataverse.Plugins.FieldService.Models;
using Microsoft.Xrm.Sdk;
using Newtonsoft.Json;
using System;

namespace Havells.Dataverse.Plugins.FieldService.Customer_APIs
{
    public class PatchSchemeCode : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion
            CustomerResult customerResult = new CustomerResult();
            string JsonResponse = "";

            try
            {
                #region Validate Params
                if (context.InputParameters.Contains("WorkOrderId") && context.InputParameters["WorkOrderId"] is string
                    && context.InputParameters.Contains("SchemeCodeId") && context.InputParameters["SchemeCodeId"] is string)
                {
                    Guid WorkOrderId = Guid.Empty;
                    bool isValidWorkOrderId = Guid.TryParse(context.InputParameters["WorkOrderId"].ToString(), out WorkOrderId);

                    Guid SchemeCodeId = Guid.Empty;
                    bool isValidSchemeCodeId = Guid.TryParse(context.InputParameters["SchemeCodeId"].ToString(), out SchemeCodeId);

                    if (!isValidWorkOrderId)
                    {
                        JsonResponse = JsonConvert.SerializeObject(new CustomerResult
                        {
                            ResultStatus = false,
                            ResultMessage = "Invalid Work Order GuId."
                        });
                        context.OutputParameters["Response"] = JsonResponse;
                        return;
                    }
                    if (!isValidSchemeCodeId)
                    {
                        JsonResponse = JsonConvert.SerializeObject(new CustomerResult
                        {
                            ResultStatus = false,
                            ResultMessage = "Invalid Scheme Code GuId."
                        });
                        context.OutputParameters["Response"] = JsonResponse;
                        return;
                    }
                    if (service != null)
                    {
                        Entity entUpdate = new Entity("msdyn_workorder", WorkOrderId);
                        entUpdate["hil_schemecode"] = new EntityReference("hil_schemeincentive", SchemeCodeId);
                        service.Update(entUpdate);
                        JsonResponse = JsonConvert.SerializeObject(new CustomerResult
                        {
                            ResultStatus = true,
                            ResultMessage = "OK"
                        });
                        context.OutputParameters["Response"] = JsonResponse;
                        return;
                    }
                    else
                    {
                        JsonResponse = JsonConvert.SerializeObject(new CustomerResult
                        {
                            ResultStatus = false,
                            ResultMessage = "D365 Service Unavailable"
                        });
                        context.OutputParameters["Response"] = JsonResponse;
                        return;
                    }
                    #endregion
                }
            }
            catch (Exception ex)
            {
                JsonResponse = JsonConvert.SerializeObject(new CustomerResult
                {
                    ResultStatus = false,
                    ResultMessage = ex.Message
                });
                context.OutputParameters["Response"] = JsonResponse;
                return;
            }
        }
    }
}
