
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsNewPlugin.SalesOrder
{
	public class PreCreate_Salesorderline : IPlugin
	{
		public static ITracingService tracingService = null;
		public void Execute(IServiceProvider serviceProvider)
		{
			#region PluginConfig
			tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
			IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
			IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
			IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
			#endregion;

			try
			{
				if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity && context.PrimaryEntityName.ToLower() == "salesorderdetail")
				{
					Entity salesOrderLine = (Entity)context.InputParameters["Target"];
					Entity orderconct = service.Retrieve(salesOrderLine.GetAttributeValue<EntityReference>("salesorderid").LogicalName, salesOrderLine.GetAttributeValue<EntityReference>("salesorderid").Id, new ColumnSet(true));
					if (salesOrderLine.Contains("productid"))
					{
						EntityReference ProductReference = (EntityReference)salesOrderLine["productid"];
						Entity productcheck = service.Retrieve("product", ProductReference.Id, new ColumnSet("hil_division", "productid"));
						tracingService.Trace("Record ID " + salesOrderLine.Id.ToString());

						if (productcheck.Contains("productid"))
						{
							Entity salesorder = service.Retrieve("salesorder", salesOrderLine.GetAttributeValue<EntityReference>("salesorderid").Id, new ColumnSet("hil_productdivision"));
							salesorder["hil_productdivision"] = productcheck.Contains("hil_division") ? new EntityReference("hil_branch", productcheck.GetAttributeValue<EntityReference>("hil_division").Id) : null;
							service.Update(salesorder);
							tracingService.Trace("Order Updated");
						}
					}

					if (orderconct.Contains("hil_ordertype") && orderconct.Contains("hil_source"))
					{
						if ((orderconct.GetAttributeValue<OptionSetValue>("hil_source").Value == 3) && (orderconct.GetAttributeValue<EntityReference>("hil_ordertype").Id.ToString() == "1f9e3353-0769-ef11-a670-0022486e4abb" || orderconct.GetAttributeValue<EntityReference>("hil_ordertype").Id.ToString() == "b8a83059-0769-ef11-a670-0022486e4abb"))
						{
							if (!salesOrderLine.Contains("hil_customerasset"))
							{
								tracingService.Trace("CustomAsset Validation Check");

								if (salesOrderLine.Contains("hil_assetmodelcode"))
								{
									tracingService.Trace("Asset model code on sales orderline");
									//	Entity orderconct = service.Retrieve(salesOrderLine.GetAttributeValue<EntityReference>("salesorderid").LogicalName, salesOrderLine.GetAttributeValue<EntityReference>("salesorderid").Id, new ColumnSet(true));
									Entity Productdata = service.Retrieve(salesOrderLine.GetAttributeValue<EntityReference>("hil_assetmodelcode").LogicalName, salesOrderLine.GetAttributeValue<EntityReference>("hil_assetmodelcode").Id, new ColumnSet(true));

									Entity Productconfigdata = service.Retrieve(Productdata.GetAttributeValue<EntityReference>("hil_materialgroup").LogicalName, Productdata.GetAttributeValue<EntityReference>("hil_materialgroup").Id, new ColumnSet(true));
									tracingService.Trace(" Assetmodelcode: materialgroup check");
									var query = new QueryExpression("hil_stagingdivisonmaterialgroupmapping");
									query.ColumnSet.AddColumn("hil_name");
									query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, Productconfigdata.GetAttributeValue<string>("name").ToString());

									EntityCollection productcon = service.RetrieveMultiple(query);

									tracingService.Trace(" productsubcateogry data  count " + productcon.Entities.Count);

									if (productcon.Entities.Count != 0)
									{
										string customerassetcheck = $@"<fetch>
														  <entity name=""msdyn_customerasset"">
															<attribute name=""msdyn_name"" />
															<attribute name=""msdyn_customerassetid"" />
															<filter>
															  <condition attribute=""msdyn_name"" operator=""eq"" value=""{salesOrderLine.GetAttributeValue<string>("hil_assetserialnumber")}"" />
															</filter>
														  </entity>
														</fetch>";

										EntityCollection customerassetcheckcoll = service.RetrieveMultiple(new FetchExpression(customerassetcheck));
										tracingService.Trace("Customer asset Collection count check" + customerassetcheckcoll.Entities.Count);

										if (customerassetcheckcoll.Entities.Count == 0)
										{

											Entity CustAsst = new Entity("msdyn_customerasset");
											CustAsst["hil_customer"] = orderconct.Contains("customerid") ? new EntityReference("contact", orderconct.GetAttributeValue<EntityReference>("customerid").Id) : null;
											CustAsst["hil_productcategory"] = Productdata.Contains("hil_division") ? new EntityReference("product", Productdata.GetAttributeValue<EntityReference>("hil_division").Id) : null;
											CustAsst["msdyn_name"] = salesOrderLine.Contains("hil_assetserialnumber") ? salesOrderLine.GetAttributeValue<string>("hil_assetserialnumber") : null;
											CustAsst["hil_productsubcategorymapping"] = new EntityReference("hil_stagingdivisonmaterialgroupmapping", productcon[0].Id);
											CustAsst["msdyn_product"] = new EntityReference("product", Productdata.Id);
											CustAsst["hil_invoiceno"] = salesOrderLine.Contains("hil_invoicenumber") ? salesOrderLine.GetAttributeValue<string>("hil_invoicenumber") : null;
											CustAsst["hil_invoicedate"] = salesOrderLine.Contains("hil_invoicedate") ? salesOrderLine.GetAttributeValue<DateTime>("hil_invoicedate") : (DateTime?)null;
											CustAsst["hil_invoicevalue"] = salesOrderLine.Contains("hil_invoicevalue") ? (decimal?)salesOrderLine.GetAttributeValue<Money>("hil_invoicevalue").Value : null; //orderline.Contains("hil_invoicevalue") ? Convert.ToDecimal(orderline.GetAttributeValue<Money>("hil_invoicevalue").Value) :(Money)null;
											CustAsst["hil_purchasedfrom"] = salesOrderLine.Contains("hil_purchasefrom") ? salesOrderLine.GetAttributeValue<string>("hil_purchasefrom") : null;
											CustAsst["hil_source"] = new OptionSetValue(4);
											CustAsst["hil_retailerpincode"] = salesOrderLine.Contains("hil_purchasefromlocation") ? salesOrderLine.GetAttributeValue<string>("hil_purchasefromlocation") : null;//"0";
											CustAsst["statuscode"] = new OptionSetValue(910590000);


											Guid recordid = service.Create(CustAsst);
											tracingService.Trace("91");
											if (recordid != null)
											{
												salesOrderLine["hil_customerasset"] = new EntityReference("msdyn_customerasset", recordid);
												tracingService.Trace("Customer Asset Created");
											}
										}

										#region elsecondition
										else
										{
											tracingService.Trace("Customer Asset found ");
											salesOrderLine["hil_customerasset"] = new EntityReference("msdyn_customerasset", customerassetcheckcoll[0].Id);

										}
										#endregion

									}
									else
									{
										throw new InvalidPluginExecutionException("Product mapping not found");
									}
								}
								else
								{
									throw new InvalidPluginExecutionException("Asset mode code not found");
								}
							}
						}
					}
				}
			}
			catch (Exception ex)
			{
				throw new InvalidPluginExecutionException("HavellsNewPlugin.SalesOrderlinePreCreate.Execute Error :  " + ex.Message); //HavellsNewPlugin.SalesOrderlinePreCreate.Execute Error
			}
		}
	}
}

