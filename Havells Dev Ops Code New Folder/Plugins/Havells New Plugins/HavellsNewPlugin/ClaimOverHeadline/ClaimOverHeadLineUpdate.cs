using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.ClaimOverHeadline
{
    public class ClaimOverHeadLineUpdate : IPlugin
    {
        public static ITracingService tracingService = null;
        Guid PerformaInvoiceID = Guid.Empty;
        public void Execute(IServiceProvider serviceProvider)
        {

            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            //foreach (var xx in context.PreEntityImages.Values)
            //    throw new InvalidPluginExecutionException("sssss " + xx.GetType());// ((EntityReference)context.PreEntityImages["im"]).Id+" / " + ((EntityReference)context.InputParameters["Target"]).Name);

            tracingService.Trace("1ss");
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {
                    //   throw new InvalidPluginExecutionException("sssss");

                    tracingService.Trace("1");
                    Entity entity = (Entity)context.InputParameters["Target"];
                    tracingService.Trace("2");
                    entity = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_performainvoice"));
                    if (entity.Contains("hil_performainvoice"))
                    {
                        PerformaInvoiceID = ((EntityReference)entity["hil_performainvoice"]).Id;

                        string _claimOverHeadLinefetch = @"<fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
                            <entity name='hil_claimoverheadline'>
                            <attribute name='hil_performainvoice' alias='hil_performainvoice' groupby='true' />
                            <attribute name='hil_amount' alias='amount' aggregate='sum' />
                            <filter type='and'>
                                <condition attribute='hil_performainvoice' operator='eq' value='" + PerformaInvoiceID + @"' />
                                <condition attribute='statecode' operator='eq' value='0' />
                            </filter>
                            </entity>
                        </fetch>";

                        EntityCollection _claimOverHeadLineColl = service.RetrieveMultiple(new FetchExpression(_claimOverHeadLinefetch));
                        tracingService.Trace("__claimOverHeadLineColl.count" + _claimOverHeadLineColl.Entities.Count);
                        if (_claimOverHeadLineColl.Entities.Count > 0)
                        {
                            tracingService.Trace("_claimOverHeadLineColl");
                            if (_claimOverHeadLineColl.Entities[0].Attributes.Contains("amount"))
                            {
                                tracingService.Trace("_claimOverHeadLineCollIF");
                                tracingService.Trace("_claimOverHeadLineColl_Amount" + ((Money)((AliasedValue)_claimOverHeadLineColl.Entities[0]["amount"]).Value).Value.ToString());
                                Money FinalAmount = ((Money)((AliasedValue)_claimOverHeadLineColl.Entities[0]["amount"]).Value);

                                tracingService.Trace("_claimOverHeadLineColl.FinalAmount" + FinalAmount.Value);
                                Entity claimheader = new Entity("hil_claimheader");
                                claimheader["hil_expenseoverheads"] = FinalAmount;
                                claimheader.Id = PerformaInvoiceID;
                                service.Update(claimheader);
                                tracingService.Trace("Tender Uploaded");
                            }
                        }
                    }
                }
                else if (context.PreEntityImages.Contains("image") && context.PreEntityImages["image"] is Entity)
                {
                    //   throw new InvalidPluginExecutionException("sssss");

                    tracingService.Trace("1");
                    //Entity entity = (Entity)context.InputParameters["Target"];
                    tracingService.Trace("2");
                    // entity = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_performainvoice"));
                    //  if (entity.Contains("hil_performainvoice"))
                    {
                        if (context.PreEntityImages.Contains("image"))
                        {
                            tracingService.Trace("3");
                            Entity preImg = ((Entity)context.PreEntityImages["image"]);
                            PerformaInvoiceID = preImg.GetAttributeValue<EntityReference>("hil_performainvoice").Id;
                        }
                        tracingService.Trace("PerformaInvoiceID " + PerformaInvoiceID);
                        //if (PerformaInvoiceID == Guid.Empty)
                        //{
                        //    PerformaInvoiceID = ((EntityReference)service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_performainvoice"))["hil_performainvoice"]).Id;
                        //}
                        string _claimOverHeadLinefetch = @"<fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
                            <entity name='hil_claimoverheadline'>
                            <attribute name='hil_performainvoice' alias='hil_performainvoice' groupby='true' />
                            <attribute name='hil_amount' alias='amount' aggregate='sum' />
                            <filter type='and'>
                                <condition attribute='hil_performainvoice' operator='eq' value='" + PerformaInvoiceID + @"' />
                                <condition attribute='statecode' operator='eq' value='0' />
                            </filter>
                            </entity>
                        </fetch>";

                        EntityCollection _claimOverHeadLineColl = service.RetrieveMultiple(new FetchExpression(_claimOverHeadLinefetch));
                        tracingService.Trace("__claimOverHeadLineColl.count" + _claimOverHeadLineColl.Entities.Count);
                        if (_claimOverHeadLineColl.Entities.Count > 0)
                        {
                            tracingService.Trace("_claimOverHeadLineColl");
                            if (_claimOverHeadLineColl.Entities[0].Attributes.Contains("amount"))
                            {
                                tracingService.Trace("_claimOverHeadLineCollIF");
                                tracingService.Trace("_claimOverHeadLineColl_Amount" + ((Money)((AliasedValue)_claimOverHeadLineColl.Entities[0]["amount"]).Value).Value.ToString());
                                Money FinalAmount = ((Money)((AliasedValue)_claimOverHeadLineColl.Entities[0]["amount"]).Value);

                                tracingService.Trace("_claimOverHeadLineColl.FinalAmount" + FinalAmount.Value);
                                Entity claimheader = new Entity("hil_claimheader");
                                claimheader["hil_expenseoverheads"] = FinalAmount;
                                claimheader.Id = PerformaInvoiceID;
                                service.Update(claimheader);
                                tracingService.Trace("Tender Uploaded");
                            }
                        }
                        else
                        {
                            Entity claimheader = new Entity("hil_claimheader");
                            claimheader["hil_expenseoverheads"] = new Money(0);
                            claimheader.Id = PerformaInvoiceID;
                            service.Update(claimheader);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace("Error " + ex.Message);
                throw new InvalidPluginExecutionException("Error " + ex.Message);
            }
        }
    }
}
