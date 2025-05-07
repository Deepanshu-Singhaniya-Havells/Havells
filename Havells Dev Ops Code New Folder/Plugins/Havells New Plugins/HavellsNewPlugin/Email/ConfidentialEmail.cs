using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Text;
//using Microsoft.Xrm.Sdk;
//using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.Email
{
    public class ConfidentialEmail : IPlugin
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
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity &&
                    context.PrimaryEntityName.ToLower() == "email" && context.Depth == 1)
                {
                    Entity entity = (Entity)context.InputParameters["Target"]; ;// service.Retrieve("email", new Guid("93c19dad-e9db-eb11-bacb-6045bd729164"), new ColumnSet(true));
                    if (!entity.Contains("statuscode"))
                    {
                        return;
                    }
                    else
                    {
                        entity = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet(true));
                        var statuscode = entity.GetAttributeValue<OptionSetValue>("statuscode").Value;
                        if (statuscode == 3) 
                        {
                            if (entity.Contains("subject") && !entity.GetAttributeValue<bool>("hil_encrypted"))
                            {
                                String subject = entity.GetAttributeValue<String>("subject");
                                tracingService.Trace("subject " + subject);
                                if (subject.ToLower().Contains("confidential"))
                                {
                                    Entity _eml = new Entity("email");
                                    _eml.Id = entity.Id;
                                    _eml["statecode"] = new OptionSetValue(0);
                                    _eml["statuscode"] = new OptionSetValue(1);
                                    service.Update(_eml);

                                    String description = entity.GetAttributeValue<String>("description");
                                    var bytes = Encoding.UTF8.GetBytes(description);

                                    var encodedString = Convert.ToBase64String(bytes);
                                    bytes = Encoding.UTF8.GetBytes("RAhgdyi" + encodedString + "Tg+DAs+==");
                                    encodedString = Convert.ToBase64String(bytes);

                                    Entity entity1 = new Entity("email");
                                    entity1.Id = entity.Id;
                                    entity1["description"] = encodedString;
                                    entity1["hil_encrypted"] = true;

                                    service.Update(entity1);


                                    _eml = new Entity("email");
                                    _eml.Id = entity.Id;
                                    _eml["statecode"] = entity["statecode"];
                                    _eml["statuscode"] = entity["statuscode"];
                                    service.Update(_eml);


                                    tracingService.Trace("subject.ToLower().Contains('confidential')");
                                }

                            }

                        }

                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(" HavellsNewPlugin.Email.ConfidentialEmail.Execute Error " + ex.Message);
            }

        }
    }
}