using System;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;

namespace HavellsNewPlugin.Actions
{
	public class AMCsalesorderandorderline_creation : IPlugin
	{
		public void Execute(IServiceProvider serviceProvider)
		{
			#region PluginConfig
			ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
			IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
			IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
			IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
			#endregion

			#region Extract value from Input parameter
			string Customerid = (string)context.InputParameters["Customerid"];
			string Addressid = (string)context.InputParameters["Addressid"];
			string Assetid = (string)context.InputParameters["Assetid"];
			string Ordertype = (string)context.InputParameters["Ordertype"];
			string Totalorderamount = (string)context.InputParameters["Totalorderamount"];
			string Receiptamount = (string)context.InputParameters["Receiptamount"];
			string TotalPayableamount = (string)context.InputParameters["TotalPayableamount"];
			string Technicianid = (string)context.InputParameters["Technicianid"];
			string Orderlinedata = (string)context.InputParameters["Orderlinedata"];

			var orderlines = JsonConvert.DeserializeObject<List<Orderline>>(Orderlinedata);
			#endregion

			#region Validation check i.e., Inputparamters

			if (string.IsNullOrEmpty(Customerid))
			{
				context.OutputParameters["Status"] = false;
				context.OutputParameters["Statusmessage"] = "Customerid is required.";
				return;
			}
			if (string.IsNullOrEmpty(Addressid))
			{
				context.OutputParameters["Status"] = false;
				context.OutputParameters["Statusmessage"] = "Addressid is required.";
				return;
			}
			if (string.IsNullOrEmpty(Assetid))
			{
				context.OutputParameters["Status"] = false;
				context.OutputParameters["Statusmessage"] = "Assetid is required.";
				return;
			}
			if (string.IsNullOrEmpty(Ordertype))
			{
				context.OutputParameters["Status"] = false;
				context.OutputParameters["Statusmessage"] = "Ordertype is required.";
				return;
			}
			if (string.IsNullOrEmpty(Totalorderamount))
			{
				context.OutputParameters["Status"] = false;
				context.OutputParameters["Statusmessage"] = "Totalorderamount is required.";
				return;
			}
			#region Not Required
			//if (string.IsNullOrEmpty(Receiptamount))
			//{
			//	context.OutputParameters["Status"] = false;
			//	context.OutputParameters["Statusmessage"] = "Receiptamount is required.";
			//	return;
			//}
			//if (string.IsNullOrEmpty(TotalPayableamount))
			//{
			//	context.OutputParameters["Status"] = false;
			//	context.OutputParameters["Statusmessage"] = "TotalPayableamount is required.";
			//	return;
			//}
			#endregion
			if (string.IsNullOrEmpty(Technicianid))
			{
				context.OutputParameters["Status"] = false;
				context.OutputParameters["Statusmessage"] = "Technicianid is required.";
				return;
			}
			if (string.IsNullOrEmpty(Orderlinedata))
			{
				context.OutputParameters["Status"] = false;
				context.OutputParameters["Statusmessage"] = "Orderlinedata is required.";
				return;
			}

			#endregion

			#region Salesorder&orderline Creation

			Entity SOEntity = new Entity("salesorder");
			SOEntity["customerid"] = new EntityReference("contact", new Guid(Customerid));
			SOEntity["msdyn_psastatusreason"] = new OptionSetValue(192350000);
			SOEntity["transactioncurrencyid"] = new EntityReference("pricelevel", new Guid("68a6a9ca-6beb-e811-a96c-000d3af05828"));
			SOEntity["ownerid"] = new EntityReference("systemuser", new Guid(Technicianid));
			SOEntity["msdyn_ordertype"] = new OptionSetValue(690970002);
			SOEntity["msdyn_account"] = new EntityReference("account", new Guid("d166ba69-65da-ec11-a7b5-6045bdad2a19")); // DummyAccount
			#region 08/11/2024 Source: TechnicianApp added  as requested by Mobile app team (Sahil Rajput)
			SOEntity["hil_source"] = new OptionSetValue(25);
			#endregion
			SOEntity["hil_serviceaddress"] = new EntityReference("hil_address", new Guid(Addressid));

			#region OrderType validation

			QueryExpression ordertypequery = new QueryExpression("hil_ordertype");
			ordertypequery.ColumnSet.AddColumns("hil_ordertype", "hil_ordertypeid", "hil_pricelist");
			ordertypequery.Criteria.AddCondition("hil_ordertypeid", ConditionOperator.Equal, Ordertype.ToString());

			EntityCollection ordertypecollection = service.RetrieveMultiple(ordertypequery);

			if (ordertypecollection.Entities.Count > 0)
			{
				string ordertype_value = ordertypecollection[0].Id.ToString();
				switch (ordertype_value)
				{
					case "1f9e3353-0769-ef11-a670-0022486e4abb":
						SOEntity["hil_ordertype"] = new EntityReference("hil_ordertype", new Guid(ordertype_value));
						break;

					case "019f761c-1669-ef11-a670-000d3a3e636d":
						SOEntity["hil_ordertype"] = new EntityReference("hil_ordertype", new Guid(ordertype_value));
						break;

					case "b8a83059-0769-ef11-a670-0022486e4abb":
						SOEntity["hil_ordertype"] = new EntityReference("hil_ordertype", new Guid(ordertype_value));
						break;

					case "22c1bc5f-0769-ef11-a670-0022486e4abb":
						SOEntity["hil_ordertype"] = new EntityReference("hil_ordertype", new Guid(ordertype_value));
						break;

					case "cad4c26b-0769-ef11-a670-0022486e4abb":
						SOEntity["hil_ordertype"] = new EntityReference("hil_ordertype", new Guid(ordertype_value));
						break;
				}
				SOEntity["pricelevelid"] = new EntityReference(ordertypecollection[0].GetAttributeValue<EntityReference>("hil_pricelist").LogicalName, ordertypecollection[0].GetAttributeValue<EntityReference>("hil_pricelist").Id);
			}
			else
			{
				context.OutputParameters["Status"] = false;
				context.OutputParameters["Statusmessage"] = "Ordertype not found.";
				return;
			}
			#endregion

			SOEntity["totallineitemamount"] = new Money(Convert.ToDecimal(Totalorderamount));
			SOEntity["hil_receiptamount"] = new Money(Convert.ToDecimal(Receiptamount));
			SOEntity["totalamount"] = new Money(Convert.ToDecimal(TotalPayableamount));

			Guid orderID = service.Create(SOEntity);

			if (orderID != null)
			{
				ColumnSet columnSalesOrder = new ColumnSet("ordernumber", "createdon", "msdyn_psastatusreason", "name");
				Entity entitySalesOrder = service.Retrieve("salesorder", orderID, columnSalesOrder);
				Entity customerasset = service.Retrieve("msdyn_customerasset", new Guid(Assetid), new ColumnSet("hil_invoicedate", "hil_invoiceno", "hil_invoicevalue", "hil_purchasedfrom", "hil_retailerpincode"));

				foreach (var line in orderlines)
				{
					Entity soitem = new Entity("salesorderdetail");
					soitem["salesorderid"] = new EntityReference("salesorder", orderID);
					soitem["productid"] = new EntityReference("product", new Guid(line.ProductID));
					soitem["hil_product"] = new EntityReference("product", new Guid(line.ProductID));
					soitem["quantity"] = Convert.ToDecimal(line.Quantity);
					soitem["priceperunit"] = new Money(Convert.ToDecimal(line.PricePerUnit));
					soitem["baseamount"] = new Money(Convert.ToDecimal(line.Amount));
					soitem["uomid"] = new EntityReference("uom", new Guid("0359d51b-d7cf-43b1-87f6-fc13a2c1dec8"));
					soitem["ownerid"] = new EntityReference("systemuser", new Guid(Technicianid));
					soitem["hil_customerasset"] = new EntityReference("msdyn_customerasset", new Guid(Assetid));
					soitem["hil_invoicedate"] = customerasset.GetAttributeValue<DateTime>("hil_invoicedate");
					soitem["hil_invoicenumber"] = customerasset.Contains("hil_invoiceno") ? customerasset.GetAttributeValue<string>("hil_invoiceno") : null;
					soitem["hil_invoicevalue"] = new Money(customerasset.GetAttributeValue<decimal>("hil_invoicevalue"));
					soitem["hil_purchasefrom"] = customerasset.Contains("hil_purchasedfrom") ? customerasset.GetAttributeValue<string>("hil_purchasedfrom") : null;
					soitem["hil_purchasefromlocation"] = customerasset.Contains("hil_retailerpincode") ? customerasset.GetAttributeValue<string>("hil_retailerpincode") : null;

					Guid orderLineID = service.Create(soitem);

					context.OutputParameters["Status"] = true;
					context.OutputParameters["OrderID"] = orderID.ToString();
					context.OutputParameters["Ordernumber"] = entitySalesOrder.GetAttributeValue<string>("name");
					context.OutputParameters["Statusmessage"] = "Order Created.";
				}
			}
			else
			{
				context.OutputParameters["Status"] = false;
				context.OutputParameters["Statusmessage"] = "Failed to create Order.";
			}
			#endregion
		}
	}
	public class Orderline
	{
		public string Amount { get; set; }
		public string PricePerUnit { get; set; }
		public string ProductID { get; set; }
		public string Quantity { get; set; }
	}

}
