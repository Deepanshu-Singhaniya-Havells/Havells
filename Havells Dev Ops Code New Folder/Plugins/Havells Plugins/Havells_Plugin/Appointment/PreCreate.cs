using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;

namespace Havells_Plugin.Appointment
{
    public class PreCreate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #region MainRegion
            try
            {
                if (context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == "phonecall"
                    //&& context.MessageName.ToUpper() == "UPDATE"
                    )
                {
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity entity = (Entity)context.InputParameters["Target"];
                    tracingService.Trace("Step 1 ");
                    if (entity.Contains("scheduledend"))
                    {
                        tracingService.Trace("Step 2 ");
                        PhoneCall appointment = entity.ToEntity<PhoneCall>();
                        if (appointment.ScheduledEnd != null)
                        {
                            tracingService.Trace("Step 3 ");
                            int startTime = 210;
                            int endTime = 870;
                            string errorMsg = "Appointment Time should be in between 9 AM to 8 PM.";

                            QueryExpression Query = new QueryExpression("hil_integrationconfiguration");
                            Query.ColumnSet = new ColumnSet("hil_appointmentstarttime", "hil_appointmentendtime", "hil_appointmenttimingsalertmessage");
                            Query.Criteria = new FilterExpression(LogicalOperator.And);
                            Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "Appointment Timings");
                            EntityCollection Found = service.RetrieveMultiple(Query);
                            foreach (Entity enConfig in Found.Entities)
                            {
                                startTime = enConfig.GetAttributeValue<Int32>("hil_appointmentstarttime");
                                endTime = enConfig.GetAttributeValue<Int32>("hil_appointmentendtime");
                                errorMsg = enConfig.GetAttributeValue<string>("hil_appointmenttimingsalertmessage");
                                tracingService.Trace("Step 4 " + startTime.ToString() + "/" + endTime.ToString());
                            }
                            if (startTime > 0)
                            {
                                startTime = startTime - 330;
                            }
                            if (endTime > 0)
                            {
                                endTime = endTime - 330;
                            }
                            int hour = appointment.ScheduledEnd.Value.Hour;
                            int minutes = (hour * 60) + appointment.ScheduledEnd.Value.Minute;

                            tracingService.Trace("Step 5 " + hour.ToString() + "/" + minutes.ToString());

                            if (minutes < startTime || minutes > endTime)
                            {
                                throw new InvalidPluginExecutionException(errorMsg);
                            }

                            #region Restrict User to book back dated Appointment {Added by Kuldeep Khare Dated: 29/Jan/2020}
                            int hourCurrent = DateTime.Now.Hour;
                            int minutesCurrent = (hourCurrent * 60) + DateTime.Now.Minute;

                            tracingService.Trace("Step 6 ");
                            DateTime date1 = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);
                            DateTime date2 = new DateTime(appointment.ScheduledEnd.Value.Year, appointment.ScheduledEnd.Value.Month, appointment.ScheduledEnd.Value.Day, appointment.ScheduledEnd.Value.Hour, appointment.ScheduledEnd.Value.Minute, appointment.ScheduledEnd.Value.Second);

                            TimeSpan ts = date2 - date1;

                            if (ts.TotalMinutes < 0.00)
                            {
                                throw new InvalidPluginExecutionException(" ***Back Date Appointment is not allowed.*** ");
                            }
                            #endregion
                        }
                    }
                    tracingService.Trace("Step Z ");
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
            #endregion
        }
    }
}
