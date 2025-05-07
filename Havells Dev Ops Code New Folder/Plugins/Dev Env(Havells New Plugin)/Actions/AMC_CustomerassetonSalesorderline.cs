using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System.Windows.Documents;

namespace HavellsNewPlugin.Actions
{
	public class AMC_CustomerassetonSalesorderline : IPlugin
	{

		public void Execute(IServiceProvider serviceProvider)
		{
			#region PluginConfig
			ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
			tracingService.Trace("18");
			IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
			IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
			IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
			tracingService.Trace("19");

			#endregion

			if (context.InputParameters.Contains("OrderID") && context.InputParameters["OrderID"] is EntityReference && context.InputParameters["OrderLineID"] == null)
			{
				tracingService.Trace("In the order condition");
				EntityReference orderId = (EntityReference)context.InputParameters["OrderID"];

				bulkcreateasset(service, tracingService, context, orderId);

			}


			else if (context.InputParameters.Contains("OrderLineID") && context.InputParameters["OrderLineID"] is EntityReference && context.InputParameters["OrderID"] == null)
			{
				tracingService.Trace("In the ordeline line condition");
				EntityReference orderLineId = (EntityReference)context.InputParameters["OrderLineID"];
				createasset(service, tracingService, context, orderLineId);
			}
			else
			{
				tracingService.Trace("Condition Not Met i.e.. Single Creation & Bulk Asset Creation");
			}
		}

		public void createasset(IOrganizationService service, ITracingService tracingService, IPluginExecutionContext context, EntityReference orderLineId)
		{
			tracingService.Trace("125 ");
			Entity orderline = service.Retrieve("salesorderdetail", orderLineId.Id, new ColumnSet("salesorderid", "productid", "hil_invoicenumber", "hil_invoicedate", "hil_invoicevalue", "hil_purchasefrom", "hil_customerasset", "hil_assetmodelcode", "hil_assetserialnumber"));
			try
			{
				if (orderline != null)
				{
					tracingService.Trace("132 ");
					Entity orderconct = service.Retrieve(orderline.GetAttributeValue<EntityReference>("salesorderid").LogicalName, orderline.GetAttributeValue<EntityReference>("salesorderid").Id, new ColumnSet(true));
					
					Entity Productdata = service.Retrieve(orderline.GetAttributeValue<EntityReference>("hil_assetmodelcode").LogicalName, orderline.GetAttributeValue<EntityReference>("hil_assetmodelcode").Id, new ColumnSet(true));

					Entity Productconfigdata = service.Retrieve(Productdata.GetAttributeValue<EntityReference>("hil_materialgroup").LogicalName, Productdata.GetAttributeValue<EntityReference>("hil_materialgroup").Id, new ColumnSet(true));
					tracingService.Trace("137 ");
					
					var query = new QueryExpression("hil_stagingdivisonmaterialgroupmapping");
					query.ColumnSet.AddColumn("hil_name");
					query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, Productconfigdata.GetAttributeValue<string>("name").ToString());

					EntityCollection productcon = service.RetrieveMultiple(query);
					tracingService.Trace("143 ");
					if (productcon.Entities.Count != 0)
					{
						string customerassetcheck = $@"<fetch>
														  <entity name=""msdyn_customerasset"">
															<attribute name=""msdyn_name"" />
															<attribute name=""msdyn_customerassetid"" />
															<filter>
															  <condition attribute=""msdyn_name"" operator=""eq"" value=""{orderline.GetAttributeValue<string>("hil_assetserialnumber")}"" />
															</filter>
														  </entity>
														</fetch>";
						EntityCollection customerassetcheckcoll = service.RetrieveMultiple(new FetchExpression(customerassetcheck));
						tracingService.Trace("156 Count : " + customerassetcheckcoll.Entities.Count);
						if (customerassetcheckcoll.Entities.Count == 0)
						{

							Entity CustAsst = new Entity("msdyn_customerasset");

							CustAsst["hil_customer"] = orderconct.Contains("customerid") ? new EntityReference("contact", orderconct.GetAttributeValue<EntityReference>("customerid").Id) : null;
							CustAsst["hil_productcategory"] = Productdata.Contains("hil_division") ? new EntityReference("product", Productdata.GetAttributeValue<EntityReference>("hil_division").Id) : null;
							CustAsst["msdyn_name"] = orderline.Contains("hil_assetserialnumber") ? orderline.GetAttributeValue<string>("hil_assetserialnumber") : null;
							CustAsst["hil_productsubcategorymapping"] = new EntityReference("hil_stagingdivisonmaterialgroupmapping", productcon[0].Id);
							CustAsst["msdyn_product"] = new EntityReference("product", Productdata.Id);
							CustAsst["hil_invoiceno"] = orderline.Contains("hil_invoicenumber") ? orderline.GetAttributeValue<string>("hil_invoicenumber") : null;
							CustAsst["hil_invoicedate"] = orderline.Contains("hil_invoicedate") ? orderline.GetAttributeValue<DateTime>("hil_invoicedate") : (DateTime?)null;
							CustAsst["hil_invoicevalue"] = orderline.Contains("hil_invoicevalue") ? (decimal?)orderline.GetAttributeValue<Money>("hil_invoicevalue").Value : null; //orderline.Contains("hil_invoicevalue") ? Convert.ToDecimal(orderline.GetAttributeValue<Money>("hil_invoicevalue").Value) :(Money)null;
							CustAsst["hil_purchasedfrom"] = orderline.Contains("hil_purchasefrom") ? orderline.GetAttributeValue<string>("hil_purchasefrom") : null;
							CustAsst["hil_source"] = new OptionSetValue(4);
							CustAsst["hil_retailerpincode"] = "0";
							CustAsst["statuscode"] = new OptionSetValue(910590000);

							Guid recordid = service.Create(CustAsst);

							if (recordid != null)
							{
								orderline["hil_customerasset"] = new EntityReference("msdyn_customerasset", recordid);
								service.Update(orderline);
						
								context.OutputParameters["Status"] = "Customer Asset Created & Updated on OrderLine";
								tracingService.Trace("Return Status: " + context.OutputParameters["Status"]);
							}
						}

						else
						{
							tracingService.Trace("189");
							orderline["hil_customerasset"] = new EntityReference("msdyn_customerasset", customerassetcheckcoll[0].Id);
							service.Update(orderline);
							context.OutputParameters["Status"] = "Customer Asset Updated on OrderLine";
							tracingService.Trace(orderline.Id.ToString());
						}

					}
					else
					{	
						context.OutputParameters["Status"] = "Product mapping not found";
					}
				}
			}

			catch (Exception ex)
			{
				tracingService.Trace(ex.Message.ToString());
			}
		}

		public void bulkcreateasset(IOrganizationService service, ITracingService tracingService, IPluginExecutionContext context, EntityReference orderId)
		{
			tracingService.Trace("193 ");
			try
			{
				string orderlinelists = $@"<fetch version=""1.0"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"">
										<entity name=""salesorderdetail"">
										<attribute name=""productid""/>
										<attribute name=""salesorderid""/>
										<attribute name=""salesorderdetailid""/>
										<attribute name=""hil_assetmodelcode""/>
										<attribute name=""hil_assetserialnumber""/>
										<attribute name=""hil_customerasset""/>
										<attribute name=""salesorderdetailname""/>
										<order attribute=""productid"" descending=""false""/>
										<link-entity name=""salesorder"" from=""salesorderid"" to=""salesorderid"" link-type=""inner"" alias=""ac"">
										<filter type=""and"">
										<condition attribute=""salesorderid"" operator=""eq"" uitype=""salesorder"" value=""{orderId.Id}""/>
										</filter>
										</link-entity>
										</entity>
										</fetch>";
				EntityCollection orderlinecoll = service.RetrieveMultiple(new FetchExpression(orderlinelists));
				tracingService.Trace("229");

				if (orderlinecoll.Entities.Count != 0)
				{
					foreach (var custass in orderlinecoll.Entities)
					{
						if (custass.GetAttributeValue<EntityReference>("hil_customerasset") == null)
						{
							Guid sorderid = custass.GetAttributeValue<EntityReference>("salesorderid").Id;
							
							Entity orderconct = service.Retrieve("salesorder", sorderid, new ColumnSet(true));
							tracingService.Trace("239");
							
							Entity Productdata = service.Retrieve(custass.GetAttributeValue<EntityReference>("hil_assetmodelcode").LogicalName, custass.GetAttributeValue<EntityReference>("hil_assetmodelcode").Id, new ColumnSet(true));
							tracingService.Trace("241");
							
							Entity Productconfigdata = service.Retrieve(Productdata.GetAttributeValue<EntityReference>("hil_materialgroup").LogicalName, Productdata.GetAttributeValue<EntityReference>("hil_materialgroup").Id, new ColumnSet(true));
							tracingService.Trace("243");

							var query = new QueryExpression("hil_stagingdivisonmaterialgroupmapping");
							query.ColumnSet.AddColumn("hil_name");
							query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, Productconfigdata.GetAttributeValue<string>("name").ToString());

							EntityCollection productcon = service.RetrieveMultiple(query);

							tracingService.Trace("250");
							tracingService.Trace("COUNT" + productcon.Entities.Count);

							if (productcon.Entities.Count != 0)
							{
								string customerassetcheck = $@"<fetch>
														  <entity name=""msdyn_customerasset"">
															<attribute name=""msdyn_name"" />
															<attribute name=""msdyn_customerassetid"" />
															<filter>
															  <condition attribute=""msdyn_name"" operator=""eq"" value=""{custass.GetAttributeValue<string>("hil_assetserialnumber")}"" />
															</filter>
														  </entity>
														</fetch>";
								EntityCollection customerassetcheckcoll = service.RetrieveMultiple(new FetchExpression(customerassetcheck));

								if (customerassetcheckcoll.Entities.Count == 0)
								{
									Entity CustAsst = new Entity("msdyn_customerasset");

									CustAsst["hil_customer"] = orderconct.Contains("customerid") ? new EntityReference("contact", orderconct.GetAttributeValue<EntityReference>("customerid").Id) : null;
									CustAsst["hil_productcategory"] = Productdata.Contains("hil_division") ? new EntityReference("product", Productdata.GetAttributeValue<EntityReference>("hil_division").Id) : null;
									CustAsst["msdyn_name"] = custass.Contains("hil_assetserialnumber") ? custass.GetAttributeValue<string>("hil_assetserialnumber") : null; //Productdata.Contains("name") ? Productdata.GetAttributeValue<string>("name") : null;
									CustAsst["hil_productsubcategorymapping"] = new EntityReference("hil_stagingdivisonmaterialgroupmapping", productcon[0].Id);
									CustAsst["msdyn_product"] = new EntityReference("product", Productdata.Id);
									CustAsst["hil_invoiceno"] = custass.Contains("hil_invoicenumber") ? custass.GetAttributeValue<string>("hil_invoicenumber") : null;
									CustAsst["hil_invoicedate"] = custass.Contains("hil_invoicedate") ? custass.GetAttributeValue<DateTime>("hil_invoicedate") : (DateTime?)null;
									CustAsst["hil_invoicevalue"] = custass.Contains("hil_invoicevalue") ? (decimal?)custass.GetAttributeValue<Money>("hil_invoicevalue").Value : null; //orderline.Contains("hil_invoicevalue") ? Convert.ToDecimal(orderline.GetAttributeValue<Money>("hil_invoicevalue").Value) :(Money)null;
									CustAsst["hil_purchasedfrom"] = custass.Contains("hil_purchasefrom") ? custass.GetAttributeValue<string>("hil_purchasefrom") : null;
									CustAsst["hil_source"] = new OptionSetValue(4);
									CustAsst["hil_retailerpincode"] = "0";
									CustAsst["statuscode"] = new OptionSetValue(910590000);

									Guid recordid = service.Create(CustAsst);
									if (recordid != null)
									{
										Entity orderline = service.Retrieve("salesorderdetail", custass.Id, new ColumnSet("hil_customerasset"));

										orderline["hil_customerasset"] = new EntityReference("msdyn_customerasset", recordid);
										service.Update(orderline);
										context.OutputParameters["Status"] = "Bulk Customer Asset Created & Updated on OrderLine";
										tracingService.Trace(orderline.Id.ToString());
									}

								}
								else
								{
									Entity orderline = service.Retrieve("salesorderdetail", custass.Id, new ColumnSet("hil_customerasset"));

									orderline["hil_customerasset"] = new EntityReference("msdyn_customerasset", customerassetcheckcoll[0].Id);
									service.Update(orderline);
									context.OutputParameters["Status"] = "Bulk Customer Asset Updated on OrderLine";
									tracingService.Trace(orderline.Id.ToString());
								}
							}
						}						
					}
				}
			}
			catch (Exception ex)
			{
				tracingService.Trace(ex.Message.ToString());
			}
		}
	}
}
