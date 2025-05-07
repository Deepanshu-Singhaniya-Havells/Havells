using Havells.Dataverse.Plugins.CommonLibs;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Runtime.Remoting.Contexts;
using System.Text;
using static Havells.Dataverse.Plugins.CommonLibs.PluginContextOps;

namespace Havells.Dataverse.Plugins.FieldService
{
    /// <summary>
    /// Plugin Assembly to extend Address Master Functionality
    /// </summary>
    public class PostOperationCreateAddress : IPlugin
    {
        private IOrganizationService _service { get; set; }
        private ITracingService _tracingService { get; set; }
        private Entity _primaryEntity { get; set; }
        public void Execute(IServiceProvider serviceProvider)
        {
            #region Execution Context Vars
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            #endregion
            try
            {
                if (PluginContextOps.ValidateContext(context, PluginMessages.Create, MDM.Address, 1))
                {
                    _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
                    IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                    _service = serviceFactory.CreateOrganizationService(context.UserId);
                    _primaryEntity = (Entity)context.InputParameters["Target"];

                    ProcessRequest(_primaryEntity, _tracingService,_service);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        private void ProcessRequest(Entity entity, ITracingService _tracingService,IOrganizationService service) {
            StringBuilder fullAddress = new StringBuilder();
            if (entity.Contains("hil_street1"))
                fullAddress.Append(entity.GetAttributeValue<string>("hil_street1") + ", ");
            if (entity.Contains("hil_street2"))
                fullAddress.Append(entity.GetAttributeValue<string>("hil_street2") + ", ");
            if (entity.Contains("hil_street3"))
                fullAddress.Append(entity.GetAttributeValue<string>("hil_street3") + ", ");
            if (entity.Contains("hil_city"))
                fullAddress.Append(service.Retrieve("hil_city", entity.GetAttributeValue<EntityReference>("hil_city").Id, new ColumnSet("hil_name")).GetAttributeValue<string>("hil_name") + ", ");
            if (entity.Contains("hil_state"))
                fullAddress.Append(service.Retrieve("hil_state", entity.GetAttributeValue<EntityReference>("hil_state").Id, new ColumnSet("hil_name")).GetAttributeValue<string>("hil_name") + ", ");
            if (entity.Contains("hil_pincode"))
                fullAddress.Append(service.Retrieve("hil_pincode", entity.GetAttributeValue<EntityReference>("hil_pincode").Id, new ColumnSet("hil_name")).GetAttributeValue<string>("hil_name") + ", ");
            
            entity["hil_fulladdress"] = fullAddress.ToString();
        }
    }
}
