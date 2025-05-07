using Havells_Plugin.SAWActivity;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;

namespace Havells_Plugin.Email
{
    public class PreCreate : IPlugin
    {
        public static ITracingService tracingService = null;
        string TenderNo = string.Empty;
        string CustomerName = string.Empty;
        string ProjectName = string.Empty;
        string _subject = string.Empty;
        string _emailSubject = string.Empty;
        Guid id = new Guid();
        public void Execute(IServiceProvider serviceProvider)
        {
            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                      && context.PrimaryEntityName.ToLower() == "email" && context.MessageName.ToUpper() == "CREATE")
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    if (entity.GetAttributeValue<EntityReference>("regardingobjectid").LogicalName != "hil_tender")
                    {
                        return;
                    }
                    else
                    {
                        EntityReference Regarding = null;
                        if (entity.Contains("subject") && entity["subject"] != null)
                        {
                            _subject = (string)entity["subject"];
                            tracingService.Trace("subject:" + _subject);
                        }
                        if (entity.Contains("regardingobjectid") && entity["regardingobjectid"] != null)
                        {
                            Regarding = (EntityReference)entity["regardingobjectid"];
                            id = Regarding.Id;
                        }
                        Entity tenderEntity = service.Retrieve(Regarding.LogicalName, Regarding.Id, new ColumnSet("hil_customerprojectname", "hil_name", "hil_customername"));
                        if (tenderEntity.Contains("hil_customerprojectname") && tenderEntity["hil_customerprojectname"] != null)
                        {
                            ProjectName = (string)tenderEntity["hil_customerprojectname"];
                        }
                        if (tenderEntity.Contains("hil_name") && tenderEntity["hil_name"] != null)
                        {
                            TenderNo = (string)tenderEntity["hil_name"];
                        }
                        if (tenderEntity.Contains("hil_customername") && tenderEntity["hil_customername"] != null)
                        {
                            EntityReference customer = (EntityReference)tenderEntity["hil_customername"];
                            Guid id = customer.Id;
                            CustomerName = customer.Name;
                        }
                        _emailSubject = TenderNo + " A/C: " + CustomerName + " Project Name: " + ProjectName;
                        tracingService.Trace("Email full subject:" + _emailSubject);
                        if (_subject.IndexOf(TenderNo) < 0)
                        {
                            entity["subject"] = _emailSubject + "-" + _subject;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}
