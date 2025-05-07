using Microsoft.Xrm.Sdk;
using System;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.Actions
{
    public class CallMaskingCustomAction : IPlugin
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
                if (context.InputParameters.Contains("TechnicianId") && context.InputParameters["TechnicianId"] is string
                        && context.InputParameters.Contains("JobId") && context.InputParameters["JobId"] is string
                        && context.InputParameters.Contains("Preference") && context.InputParameters["Preference"] is int
                        && context.Depth == 1)
                {
                    try
                    {
                        string TechnicianId = (string)context.InputParameters["TechnicianId"];
                        string JobId = (string)context.InputParameters["JobId"];
                        int Preference = (int)context.InputParameters["Preference"];

                        Entity technician = service.Retrieve("systemuser", new Guid(TechnicianId), new ColumnSet("mobilephone"));
                        string PatchingNumber = technician.Contains("mobilephone") ? technician.GetAttributeValue<string>("mobilephone") : null;
                        if (PatchingNumber == null)
                        {
                            context.OutputParameters["Result"] = "Technician Number not found ";
                            return;
                        }
                        ReqCall req = new ReqCall
                        {
                            MobileNumber = PatchingNumber.Substring(Math.Max(0, PatchingNumber.Length - 10)),
                            JobId = new Guid(JobId),
                            Preference = Preference,
                            TechnicianId = new Guid(TechnicianId)
                        };
                        Guid _phoneCallId = CreateCallInteraction(service, req);
                        if (_phoneCallId != Guid.Empty)
                        {
                            context.OutputParameters["Result"] = "Success";
                        }
                        else
                        {
                            context.OutputParameters["Result"] = "Failure";
                        }
                    }
                    catch (Exception ex)
                    {
                        context.OutputParameters["Result"] = "D365 Internal Server Error : " + ex.Message;
                    }
                }
            }
            catch (Exception ex)
            {
                context.OutputParameters["Result"] = "D365 Internal Server Error : " + ex.Message;
            }
        }
        public static Guid CreateCallInteraction(IOrganizationService service, ReqCall req)
        {
            Guid res = Guid.Empty;
            try
            {
                Entity phoneCall = new Entity("phonecall");
                phoneCall["hil_callingnumber"] = req.MobileNumber;
                phoneCall["hil_contactpreference"] = new OptionSetValue(req.Preference);
                phoneCall["regardingobjectid"] = new EntityReference("msdyn_workorder", req.JobId);
                phoneCall["ownerid"] = new EntityReference("systemuser", req.TechnicianId);
                res = service.Create(phoneCall);
            }
            catch (Exception)
            {
            }
            return res;
        }
    }
    public class ReqCall
    {
        public string MobileNumber { get; set; }
        public int Preference { get; set; }
        public Guid JobId { get; set; }
        public Guid TechnicianId { get; set; }
    }
}
