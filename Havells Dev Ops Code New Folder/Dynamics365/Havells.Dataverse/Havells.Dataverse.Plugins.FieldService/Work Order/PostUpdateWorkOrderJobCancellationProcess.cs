using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace Havells.Dataverse.Plugins.FieldService.Work_Order
{
    public class PostUpdateWorkOrderJobCancellationProcess : IPlugin
    {
        private static readonly Guid RequestForCancellation = new Guid("73edaa5e-0211-f011-9989-7ced8d26d424"); //Request For Cancellation
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                && context.PrimaryEntityName.ToLower() == "msdyn_workorder" && (context.MessageName.ToUpper() == "UPDATE"))
            {
                try
                {
                    Entity entity = (Entity)context.InputParameters["Target"];

                    if (entity.Contains("hil_jobcancelreason"))
                    {
                        try
                        {
                            OptionSetValue _jobcancelreason = entity.GetAttributeValue<OptionSetValue>("hil_jobcancelreason");
                            if (_jobcancelreason != null && _jobcancelreason.Value == 7) //Product Working Satisfactorily
                            {
                                Entity jobRef = new Entity("msdyn_workorder", entity.Id);
                                jobRef["msdyn_substatus"] = new EntityReference("msdyn_workordersubstatus", RequestForCancellation); //Request For Cancellation
                                service.Update(jobRef);

                                Entity WaCta = new Entity("hil_wacta");
                                WaCta["hil_watriggername"] = "cancel_2cta_h8";
                                WaCta["hil_wastatusreason"] = new OptionSetValue(1);
                                WaCta["hil_jobid"] = new EntityReference("msdyn_workorder", jobRef.Id);
                                service.Create(WaCta);
                            }
                            if (_jobcancelreason != null && (_jobcancelreason.Value == 3 || _jobcancelreason.Value == 2)) //Rescheduled Appointment OR Site not Ready
                            {
                                Entity jobRef = new Entity("msdyn_workorder", entity.Id);
                                jobRef["msdyn_substatus"] = new EntityReference("msdyn_workordersubstatus", RequestForCancellation); //Request For Cancellation
                                service.Update(jobRef);

                                Entity WaCta = new Entity("hil_wacta");
                                WaCta["hil_watriggername"] = "cancel_3cta_jo";
                                WaCta["hil_wastatusreason"] = new OptionSetValue(1);
                                WaCta["hil_jobid"] = new EntityReference("msdyn_workorder", jobRef.Id);
                                service.Create(WaCta);
                            }

                            if (_jobcancelreason != null && _jobcancelreason.Value == 4) // Duplicate request
                            {
                                EntityReference DuplicateJobRef = entity.Contains("hil_newjobid") ? entity.GetAttributeValue<EntityReference>("hil_newjobid") : null;
                                if (DuplicateJobRef != null)
                                {
                                    Entity jobRef = new Entity("msdyn_workorder", entity.Id);
                                    jobRef["msdyn_substatus"] = new EntityReference("msdyn_workordersubstatus", RequestForCancellation); //Request For Cancellation
                                    service.Update(jobRef);

                                    Entity WaCta = new Entity("hil_wacta");
                                    WaCta["hil_watriggername"] = "cancel_nocta_jp";
                                    WaCta["hil_wastatusreason"] = new OptionSetValue(1);
                                    WaCta["hil_jobid"] = new EntityReference("msdyn_workorder", jobRef.Id);
                                    service.Create(WaCta);
                                }
                            }

                            if (_jobcancelreason != null && (_jobcancelreason.Value == 5 || _jobcancelreason.Value == 6 || _jobcancelreason.Value == 8)) //Product Replaced by Dealer
                            {
                                Entity jobRef = new Entity("msdyn_workorder", entity.Id);
                                jobRef["msdyn_substatus"] = new EntityReference("msdyn_workordersubstatus", RequestForCancellation); //Request For Cancellation
                                service.Update(jobRef);
                            }

                            if (_jobcancelreason != null && _jobcancelreason.Value == 9) // Bill or Invoice Not Available	
                            {
                                _tracingService.Trace("inside bill not available");
                                if (!entity.Contains("msdyn_customerasset"))
                                {
                                    _tracingService.Trace("inside because asset is null");
                                    Entity jobRef = new Entity("msdyn_workorder", entity.Id);
                                    jobRef["msdyn_substatus"] = new EntityReference("msdyn_workordersubstatus", RequestForCancellation); //Request For Cancellation
                                    service.Update(jobRef);

                                    Entity WaCta = new Entity("hil_wacta");
                                    WaCta["hil_watriggername"] = "invoice_2cta_68";
                                    WaCta["hil_wastatusreason"] = new OptionSetValue(1);
                                    WaCta["hil_jobid"] = new EntityReference("msdyn_workorder", jobRef.Id);
                                    service.Create(WaCta);
                                    _tracingService.Trace("CTA created");
                                }
                            }

                            if (_jobcancelreason != null && _jobcancelreason.Value == 10) //Estimate not approved
                            {
                                Entity jobRef = new Entity("msdyn_workorder", entity.Id);
                                jobRef["msdyn_substatus"] = new EntityReference("msdyn_workordersubstatus", RequestForCancellation); //Request For Cancellation
                                service.Update(jobRef);

                                Entity WaCta = new Entity("hil_wacta");
                                WaCta["hil_watriggername"] = "est_na_2cta";
                                WaCta["hil_wastatusreason"] = new OptionSetValue(1);
                                WaCta["hil_jobid"] = new EntityReference("msdyn_workorder", jobRef.Id);
                                service.Create(WaCta);
                            }
                        }
                        catch (Exception ex)
                        {
                            throw new InvalidPluginExecutionException("PostUpdateWorkOrderJobCancellationProcess Plugin::" + ex.Message);
                        }
                    }

                    if (entity.Contains("hil_reopenbyconsumer"))
                    {
                        bool ReopenByConsumer = entity.GetAttributeValue<bool>("hil_reopenbyconsumer");
                        if (ReopenByConsumer)
                        {
                            Entity jobRef = new Entity("msdyn_workorder", entity.Id);
                            //jobRef["hil_jobcancelreason"] = new OptionSetValue(-1);
                            jobRef["msdyn_substatus"] = new EntityReference("msdyn_workordersubstatus", new Guid("2b27fa6c-fa0f-e911-a94e-000d3af060a1")); //Work Initiated
                            try
                            {
                                service.Update(jobRef);
                            }
                            catch (Exception ex)
                            {
                                throw new InvalidPluginExecutionException(ex.Message);
                            }

                            string fetchPermittedCount = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='hil_wacta'>
                            <attribute name='hil_wactaid'/>
                            <filter type='and'>
                            <condition attribute='statecode' operator='eq' value='0'/>
                            <condition attribute='hil_jobid' operator='eq' value='{jobRef.Id}'/>
                            <condition attribute='hil_wastatusreason' operator='eq' value='1'/>
                            </filter>
                            </entity>
                            </fetch>";
                            EntityCollection permittedCountColl = service.RetrieveMultiple(new FetchExpression(fetchPermittedCount));
                            if (permittedCountColl.Entities.Count > 0)
                            {
                                foreach (Entity ent in permittedCountColl.Entities)
                                {
                                    ent["hil_wastatusreason"] = new OptionSetValue(2);
                                    try
                                    {
                                        service.Update(ent);
                                    }
                                    catch (Exception ex)
                                    {
                                        throw new InvalidPluginExecutionException(ex.Message);
                                    }
                                }
                            }
                        }
                    }

                    if (entity.Contains("msdyn_substatus"))
                    {
                        EntityReference _jobSubStatus = entity.GetAttributeValue<EntityReference>("msdyn_substatus");
                        string jobStatusClosed = "1727fa6c-fa0f-e911-a94e-000d3af060a1"; //Closed - Sub-status 
                        string jobStatusCancelled = "1527fa6c-fa0f-e911-a94e-000d3af060a1"; //Cancelled - Sub-status

                        if (_jobSubStatus.Id.Equals(new Guid(jobStatusClosed)) || _jobSubStatus.Id.Equals(new Guid(jobStatusCancelled)))
                        {
                            string fetchPermittedCount = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='hil_wacta'>
                            <attribute name='hil_wactaid'/>
                            <filter type='and'>
                            <condition attribute='statecode' operator='eq' value='0'/>
                            <condition attribute='hil_jobid' operator='eq' value='{entity.Id}'/>
                            <condition attribute='hil_wastatusreason' operator='eq' value='1'/>
                            </filter>
                            </entity>
                            </fetch>";
                            EntityCollection permittedCountColl = service.RetrieveMultiple(new FetchExpression(fetchPermittedCount));
                            if (permittedCountColl.Entities.Count > 0)
                            {
                                foreach (Entity ent in permittedCountColl.Entities)
                                {
                                    ent["hil_wastatusreason"] = new OptionSetValue(2);//Closed
                                    service.Update(ent);
                                }
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
}
