using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.StagingPricingMapping
{
    public class stagingpricingmappingpostcreate : IPlugin
    {
        public static ITracingService tracingService = null;
        public static readonly string pricelevel = "5688c4d2-bf70-ef11-a670-6045bdced00e";

        public void Execute(IServiceProvider serviceProvider)
        {
            #region PluginConfig
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity && (context.MessageName.ToUpper() == "CREATE") && (context.Stage.ToString() == "40"))
            {
                Entity spmapp = (Entity)context.InputParameters["Target"];
                tracingService.Trace("Started..!");
                try
                {
                    DateTime startDate = spmapp.GetAttributeValue<DateTime>("hil_datestart");
                    DateTime endDate = spmapp.GetAttributeValue<DateTime>("hil_dateend");
                    DateTime currentDate = DateTime.Now;

                    if (currentDate >= startDate && currentDate <= endDate)
                    {
                        if (!string.IsNullOrEmpty(spmapp.GetAttributeValue<string>("hil_name")))
                        {
                            string fetch = $@"<fetch>
										  <entity name='product'>
											<attribute name='name' />
											<attribute name='hil_amount' />
											<attribute name='productid' />
											<attribute name='productnumber' />
											<attribute name='hil_hierarchylevel' />
											<filter>
											  <condition attribute='name' operator='eq' value='{spmapp.GetAttributeValue<string>("hil_name")}' />
											</filter>
										  </entity>
										</fetch>";
                            EntityCollection Spmappcoll = service.RetrieveMultiple(new FetchExpression(fetch));
							OptionSetValue _hierarchyLevel = null;
                            if (Spmappcoll.Entities.Count != 0)
                            {
								_hierarchyLevel = Spmappcoll.Entities[0].GetAttributeValue<OptionSetValue>("hil_hierarchylevel");
								if (_hierarchyLevel.Value == 910590002) //A La Carte 
								{
									Entity product = new Entity("product");
									product["productid"] = product["productid"] = new Guid(Spmappcoll.Entities[0].Id.ToString());
									product["pricelevelid"] = new EntityReference("pricelevel", new Guid(pricelevel));
									service.Update(product);

									string priceitem = $@"<fetch>
												  <entity name='productpricelevel'>
													<attribute name='amount' />
													<attribute name='hil_discount' />
													<attribute name='productid' />
													<attribute name='pricelevelid' />
													<filter>
													  <condition attribute='productid' operator='eq' value='{Spmappcoll.Entities[0].Id}' />
													</filter>
													<filter>
													  <condition attribute='pricelevelid' operator='eq' value='{pricelevel}' />
													</filter>
												  </entity>
												</fetch>";
									EntityCollection pplcoll = service.RetrieveMultiple(new FetchExpression(priceitem));
									if (pplcoll.Entities.Count > 0)
									{
										if (spmapp.GetAttributeValue<bool>("hil_type") == false)
										{
											Entity ppl = new Entity("productpricelevel");
											ppl["productpricelevelid"] = pplcoll[0].Id;
											ppl["hil_discount"] = new Money(Math.Abs(spmapp.Contains("hil_price") ? spmapp.GetAttributeValue<int>("hil_price") : 0));
											service.Update(ppl);
										}
										else
										{
											Entity ppl = new Entity("productpricelevel");
											ppl["productpricelevelid"] = pplcoll[0].Id;
											ppl["amount"] = new Money(spmapp.Contains("hil_price") ? spmapp.GetAttributeValue<int>("hil_price") : 0);
											service.Update(ppl);
										}
									}
									else
									{
										Entity ppl = new Entity("productpricelevel");
										if (spmapp.GetAttributeValue<bool>("hil_type") == false)
										{
											ppl["hil_discount"] = new Money(Math.Abs(spmapp.Contains("hil_price") ? spmapp.GetAttributeValue<int>("hil_price") : 0));
										}
										else if (spmapp.GetAttributeValue<bool>("hil_type") == true)
										{
											ppl["amount"] = new Money(spmapp.Contains("hil_price") ? spmapp.GetAttributeValue<int>("hil_price") : 0);
										}
										ppl["pricelevelid"] = new EntityReference("pricelevel", new Guid("5688c4d2-bf70-ef11-a670-6045bdced00e"));
										ppl["productid"] = new EntityReference("product", Spmappcoll.Entities[0].Id);
										ppl["uomid"] = new EntityReference("uom", new Guid("0359d51b-d7cf-43b1-87f6-fc13a2c1dec8"));
										ppl["transactioncurrencyid"] = new EntityReference("transactioncurrency", new Guid("68a6a9ca-6beb-e811-a96c-000d3af05828"));
										ppl["pricingmethodcode"] = new OptionSetValue(1);
										ppl["quantitysellingcode"] = new OptionSetValue(1);
										service.Create(ppl);
									}
								}
							}
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException("HavellsNewPlugin.StagingPricingMappingPostCreate.Execute Error :  " + ex.Message);
                }
            }
        }
    }
}
