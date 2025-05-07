using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.Payment_Receipt
{
	public class Paymentreceipt_precreate : IPlugin
	{
		public static ITracingService tracingService = null;
		IPluginExecutionContext context;
		public void Execute(IServiceProvider serviceProvider)
		{
			tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
			context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
			IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
			IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
			try
			{
				if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
					&& context.PrimaryEntityName.ToLower() == "hil_paymentreceipt" && context.MessageName.ToUpper() == "CREATE")
				{
					Entity entity = (Entity)context.InputParameters["Target"];
					tracingService.Trace("Start...");
					string prefix = "D365";
					string Counter = GetCounter(service, entity);
					tracingService.Trace("Counter is " + Counter);
					//EntityReference EntityOrderDept = entity.GetAttributeValue<EntityReference>("salesorder");
					Entity EntityOrder = service.Retrieve(entity.GetAttributeValue<EntityReference>("hil_orderid").LogicalName, entity.GetAttributeValue<EntityReference>("hil_orderid").Id, new ColumnSet("name"));
					string OrderID = EntityOrder.Contains("name") ? EntityOrder.GetAttributeValue<string>("name") : "2024110900001";
					tracingService.Trace("OrderID is " + OrderID);

					string TransactionID = prefix + OrderID + Counter;
					entity["hil_transactionid"] = TransactionID;

				}
			}
			catch (Exception ex)
			{
				throw new InvalidPluginExecutionException("HavellsNewPlugin.Payment_Receipt.Paymentreceipt_precreate : " + ex.Message);
			}
		}
		private string GetCounter(IOrganizationService service, Entity entity)
		{
			int Counter = 001;
			QueryExpression queryExp = new QueryExpression(entity.LogicalName);
			queryExp.ColumnSet = new ColumnSet("hil_transactionid");
			queryExp.AddOrder("createdon", OrderType.Descending);
			queryExp.TopCount = 1;
			EntityCollection entColl = service.RetrieveMultiple(queryExp);
			tracingService.Trace("Entity Count " + entColl.Entities.Count);

			if (entColl.Entities.Count > 0)
			{
				tracingService.Trace("Fount Payment Receipt...");
				string transactionID = entColl.Entities[0].Contains("hil_transactionid") ? entColl.Entities[0].GetAttributeValue<string>("hil_transactionid") : null;
				tracingService.Trace("Transaction ID " + transactionID);
				if (transactionID != null)
				{
					// example string == D365AMC2024080700001001
					string substring1 = transactionID.Substring(transactionID.Length - 3); // "OO1" - to get last 3 digits
					Counter = Convert.ToInt32(substring1) + 1;
					tracingService.Trace("Counter " + Counter);
				}
			}
			tracingService.Trace("Counter " + Counter);
			return Counter.ToString("D3");
		}
	}
}