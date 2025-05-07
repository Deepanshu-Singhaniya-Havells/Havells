using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.Job_Request
{
	public class Workorderrequest_create : IPlugin
	{
		private readonly int callCenterSource = 2;
		private readonly int bulkJobUploaderSource = 14;
		public static ITracingService tracingService = null;
		public void Execute(IServiceProvider serviceProvider)
		{
			#region PluginConfig
			tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
			IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
			IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
			IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);


			#endregion
			if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity && (context.MessageName.ToUpper() == "CREATE"))
			{
				
				Entity workOrderRequest = (Entity)context.InputParameters["Target"];
				try
				{
					int source = workOrderRequest.Contains("hil_source") ? workOrderRequest.GetAttributeValue<OptionSetValue>("hil_source").Value : -1;
					int mode = workOrderRequest.Contains("hil_mode") ? workOrderRequest.GetAttributeValue<OptionSetValue>("hil_mode").Value : -1;

					if (source == -1 && mode == -1) 
					{ 
						throw new InvalidPluginExecutionException("Mode Should be filled by "+ "'Bulk Upload'");
					}

					if (source == callCenterSource && mode == 1  )
					{
						Callstaging(service,tracingService ,workOrderRequest);
					}
					else if ( mode == 2 && source == -1)
					{
						if (workOrderRequest.GetAttributeValue<string>("hil_name") != null)

						{
							tracingService.Trace("46");
							BulkJobUploader(service, tracingService, workOrderRequest);
						}

						else
						{
							throw new InvalidPluginExecutionException("Fill all required field in the excel");
						}

					}
				}
				catch (Exception ex)
				{
					throw new InvalidPluginExecutionException(ex.Message);
				}
			}
		}
		private static void Callstaging(IOrganizationService service, ITracingService tracingService, Entity workOrderRequest)
		{
			try
			{
				#region Customer_Creation

				Entity entConsumer = new Entity("contact");
				entConsumer["mobilephone"] = workOrderRequest.GetAttributeValue<string>("hil_registerednumber");
				entConsumer["address1_telephone2"] = workOrderRequest.GetAttributeValue<string>("hil_name");
				entConsumer["firstname"] = workOrderRequest.GetAttributeValue<string>("hil_firstname");
				entConsumer["lastname"] = workOrderRequest.Contains("hil_lastname") ? workOrderRequest.GetAttributeValue<string>("hil_lastname") : null;
				entConsumer["emailaddress1"] = workOrderRequest.Contains("hil_email") ? workOrderRequest.GetAttributeValue<string>("hil_email") : null;
				entConsumer["address1_telephone3"] = workOrderRequest.Contains("hil_alternativenumber") ? workOrderRequest.GetAttributeValue<string>("hil_alternativenumber") : null;
				if (workOrderRequest.Contains("hil_gender"))
				{
					int gendervalue = workOrderRequest.GetAttributeValue<OptionSetValue>("hil_gender").Value;
					tracingService.Trace("Gender value " + gendervalue);
					if (gendervalue == 1)
					{
						entConsumer["hil_gender"] = true;
					}
					else if (gendervalue == 2)
					{
						entConsumer["hil_gender"] = false;
					}
					else if (gendervalue == 3)
					{
						entConsumer["hil_gender"] = null;
					}
				}
				entConsumer["hil_consumersource"] = new OptionSetValue(1);
				entConsumer["hil_subscribeformessagingservice"] = true;
				//entConsumer["hil_preferredlanguageforcommunication"] = new EntityReference("hil_preferredlanguageforcommunication", new Guid("d825675f-37db-ec11-a7b5-6045bdad294c"));

				Guid contactId = service.Create(entConsumer);
				tracingService.Trace("Consumer_ID" + contactId);
				#endregion

				#region Address_Creation
				Entity address = new Entity("hil_address");
				address["hil_customer"] = new EntityReference("contact", contactId);
				address["hil_street1"] = workOrderRequest.GetAttributeValue<string>("hil_addressline1");
				address["hil_street2"] = workOrderRequest.GetAttributeValue<string>("hil_addressline2");
				address["hil_addresstype"] = new OptionSetValue(1);
				address["hil_businessgeo"] = workOrderRequest.GetAttributeValue<EntityReference>("hil_pincode");

				Guid Addressid = service.Create(address);

				tracingService.Trace("Address_ID" + Addressid);
				#endregion

				#region Workorder_Creation

				//Entity contactdata = service.Retrieve("contact", contactId, new ColumnSet("mobilephone", "address1_telephone2", "address1_telephone3", "emailaddress1"));

				Entity Workorder = new Entity("msdyn_workorder");
				Workorder["hil_customerref"] = new EntityReference("contact", contactId);
				Workorder["hil_customername"] = workOrderRequest.GetAttributeValue<string>("hil_firstname") + " " + workOrderRequest.GetAttributeValue<string>("hil_lastname");
				Workorder["hil_mobilenumber"] = workOrderRequest.GetAttributeValue<string>("hil_registerednumber");
				Workorder["hil_callingnumber"] = workOrderRequest.GetAttributeValue<string>("hil_name");
				Workorder["hil_alternate"] = workOrderRequest.Contains("hil_alternativenumber") ? workOrderRequest.GetAttributeValue<string>("hil_alternativenumber") : null;
				Workorder["hil_email"] = workOrderRequest.Contains("hil_email") ? workOrderRequest.GetAttributeValue<string>("hil_email") : null;
				Workorder["hil_preferredtime"] = workOrderRequest.GetAttributeValue<OptionSetValue>("hil_preferreddaytime");
				Workorder["hil_preferreddate"] = workOrderRequest.GetAttributeValue<DateTime>("hil_preferreddate");
				Workorder["hil_address"] = new EntityReference("hil_address", Addressid);
				Workorder["hil_productcategory"] = workOrderRequest.GetAttributeValue<EntityReference>("hil_productcategory");
				Workorder["hil_productcatsubcatmapping"] = workOrderRequest.GetAttributeValue<EntityReference>("hil_productsubcategory");
				Workorder["hil_consumertype"] = new EntityReference("hil_consumertype", new Guid("484897de-2abd-e911-a957-000d3af0677f"));
				Workorder["hil_consumercategory"] = new EntityReference("hil_consumercategory", new Guid("baa52f5e-2bbd-e911-a957-000d3af0677f"));
				Workorder["hil_natureofcomplaint"] = workOrderRequest.GetAttributeValue<EntityReference>("hil_natureofcomplaint");
				Workorder["hil_callsubtype"] = workOrderRequest.GetAttributeValue<EntityReference>("hil_callsubtype");
				Workorder["hil_quantity"] = workOrderRequest.GetAttributeValue<int>("hil_quantity"); ;
				Workorder["hil_callertype"] = workOrderRequest.GetAttributeValue<OptionSetValue>("hil_callertype");
				Workorder["hil_newserialnumber"] = workOrderRequest.GetAttributeValue<string>("hil_newserialnumber");
				Workorder["msdyn_serviceaccount"] = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43"));
				Workorder["msdyn_billingaccount"] = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43"));
				Workorder["hil_customercomplaintdescription"] = workOrderRequest.Contains("hil_customerdescriptionremarks") ? workOrderRequest.GetAttributeValue<string>("hil_customerdescriptionremarks") : null;
				Workorder["hil_emergencyremarks"] = workOrderRequest.Contains("hil_emergencyremarks") ? workOrderRequest.GetAttributeValue<string>("hil_emergencyremarks") : null;
				Workorder["hil_purchasedate"] = workOrderRequest.Contains("hil_purchasedate") ? workOrderRequest.GetAttributeValue<DateTime>("hil_purchasedate") : (DateTime?)null;
				Workorder["hil_emergencycall"] = workOrderRequest.Contains("hil_emergencycall") ? workOrderRequest.GetAttributeValue<OptionSetValue>("hil_emergencycall") : null;

				Guid workorderid = service.Create(Workorder);
				tracingService.Trace("Job_ID" + workorderid);
				#endregion

				if (workorderid != Guid.Empty)
				{
					Entity jobRequest = new Entity(workOrderRequest.LogicalName, workOrderRequest.Id);
					jobRequest["hil_jobid"] = new EntityReference("msdyn_workorder", workorderid);
					service.Update(jobRequest);
				}
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}
		private static Guid Contactcreation(IOrganizationService service, ITracingService tracingService, Entity workOrderRequest)
		{

			QueryExpression query = new QueryExpression("contact");

			query.ColumnSet.AddColumns("address1_telephone2", "address1_telephone3", "emailaddress1", "firstname", "hil_gender", "lastname", "mobilephone");
			query.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, workOrderRequest.GetAttributeValue<string>("hil_registerednumber"));
			EntityCollection contactcheck = service.RetrieveMultiple(query);

			if (contactcheck.Entities.Count == 0)
			{

				#region Customer_Creation
				Entity entConsumer = new Entity("contact");
				entConsumer["mobilephone"] = workOrderRequest.GetAttributeValue<string>("hil_registerednumber");
				entConsumer["address1_telephone2"] = workOrderRequest.GetAttributeValue<string>("hil_name");
				entConsumer["firstname"] = workOrderRequest.GetAttributeValue<string>("hil_firstname");
				entConsumer["lastname"] = workOrderRequest.Contains("hil_lastname") ? workOrderRequest.GetAttributeValue<string>("hil_lastname") : null;
				entConsumer["emailaddress1"] = workOrderRequest.Contains("hil_email") ? workOrderRequest.GetAttributeValue<string>("hil_email") : null;
				entConsumer["address1_telephone3"] = workOrderRequest.Contains("hil_alternativenumber") ? workOrderRequest.GetAttributeValue<string>("hil_alternativenumber") : null;
				if (workOrderRequest.Contains("hil_gender"))
				{
					int gendervalue = workOrderRequest.GetAttributeValue<OptionSetValue>("hil_gender").Value;
					tracingService.Trace("Gender value " + gendervalue);
					if (gendervalue == 1)
					{
						entConsumer["hil_gender"] = true;

					}
					else if (gendervalue == 2)
					{
						entConsumer["hil_gender"] = false;

					}
					else if (gendervalue == 3)
					{
						entConsumer["hil_gender"] = null;
					}
				}
				entConsumer["hil_consumersource"] = new OptionSetValue(1);
				entConsumer["hil_subscribeformessagingservice"] = true;
				entConsumer["hil_preferredlanguageforcommunication"] = new EntityReference("hil_preferredlanguageforcommunication", new Guid("d825675f-37db-ec11-a7b5-6045bdad294c"));

				Guid contactId = service.Create(entConsumer);
				tracingService.Trace("Consumer_ID" + contactId);
				return contactId;

				#endregion
			}
			else
			{
				return contactcheck[0].Id;
			}
		}

		private static Guid Addresscretion(IOrganizationService service, ITracingService tracingService, Entity workOrderRequest)
		{

			Guid contactid = Contactcreation(service, tracingService, workOrderRequest);

			var query = new QueryExpression("hil_address");
			query.ColumnSet.AddColumns("hil_customer", "hil_street1", "hil_street2", "hil_businessgeo");
			query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, contactid);
			query.Criteria.AddCondition("hil_businessgeo", ConditionOperator.Equal, workOrderRequest.GetAttributeValue<EntityReference>("hil_pincode").Id);

			EntityCollection addresscheck = service.RetrieveMultiple(query);

			if (addresscheck.Entities.Count == 0)
			{
				Entity address = new Entity("hil_address");
				address["hil_customer"] = new EntityReference("contact", contactid);
				address["hil_street1"] = workOrderRequest.GetAttributeValue<string>("hil_addressline1");
				address["hil_street2"] = workOrderRequest.GetAttributeValue<string>("hil_addressline2");
				address["hil_addresstype"] = new OptionSetValue(1);
				address["hil_businessgeo"] = workOrderRequest.GetAttributeValue<EntityReference>("hil_pincode");

				Guid Addressid = service.Create(address);
				tracingService.Trace("Address Created ");
				return Addressid;
			}
			else
			{
				return addresscheck[0].Id;
			}
		}
		private static void Jobcretion(IOrganizationService service, ITracingService tracingService, Entity workOrderRequest)
		{
			#region Workorder_Creation

			Guid contactId = Contactcreation(service, tracingService, workOrderRequest);
			Guid addressId = Addresscretion(service, tracingService, workOrderRequest);

			Guid query_And3_hil_productcat = workOrderRequest.GetAttributeValue<EntityReference>("hil_productcategory").Id;

			var query = new QueryExpression("msdyn_workorder");
			query.ColumnSet =new ColumnSet("msdyn_name");

			query = new QueryExpression("msdyn_workorder");
			query.ColumnSet = new ColumnSet("msdyn_name");
			query.Criteria = new FilterExpression(LogicalOperator.And);
			query.Criteria.AddCondition(new ConditionExpression("hil_customerref", ConditionOperator.Equal, contactId));
			query.Criteria.AddCondition(new ConditionExpression("hil_productcategory", ConditionOperator.Equal, query_And3_hil_productcat));
			query.Criteria.AddCondition(new ConditionExpression("createdon", ConditionOperator.LastXDays, 30));
			query.Criteria.AddCondition(new ConditionExpression("msdyn_substatus", ConditionOperator.NotIn, new object[] { new Guid("1527FA6C-FA0F-E911-A94E-000D3AF060A1"), new Guid("1727FA6C-FA0F-E911-A94E-000D3AF060A1"), new Guid("6C8F2123-5106-EA11-A811-000D3AF057DD"), new Guid("2927FA6C-FA0F-E911-A94E-000D3AF060A1"), new Guid("7E85074C-9C54-E911-A951-000D3AF0677F") }));
			query.AddOrder("createdon", OrderType.Descending);

			EntityCollection jobcheck = service.RetrieveMultiple(query);

			if (jobcheck.Entities.Count == 0)
			{
				Entity Contactdata = service.Retrieve("contact", contactId, new ColumnSet("mobilephone", "address1_telephone2", "address1_telephone3", "emailaddress1"));
				Entity AddressSalesoffice = service.Retrieve("hil_address", addressId, new ColumnSet("hil_salesoffice"));

				Entity Workorder = new Entity("msdyn_workorder");
				Workorder["hil_salesoffice"] = AddressSalesoffice.Contains("hil_salesoffice") ? AddressSalesoffice.GetAttributeValue<EntityReference>("hil_salesoffice") : null;
				Workorder["hil_sourceofjob"] = new OptionSetValue(14);
				Workorder["hil_customerref"] = new EntityReference("contact", contactId);
				Workorder["hil_customername"] = workOrderRequest.GetAttributeValue<string>("hil_firstname") + " " + workOrderRequest.GetAttributeValue<string>("hil_lastname");
				Workorder["hil_mobilenumber"] = Contactdata.Contains("mobilephone") ? Contactdata.GetAttributeValue<string>("mobilephone") : null;
				Workorder["hil_callingnumber"] = Contactdata.Contains("address1_telephone2") ? Contactdata.GetAttributeValue<string>("address1_telephone2") : null;
				Workorder["hil_alternate"] = Contactdata.Contains("address1_telephone3") ? Contactdata.GetAttributeValue<string>("address1_telephone3") : null;
				Workorder["hil_email"] = Contactdata.Contains("emailaddress1") ? Contactdata.GetAttributeValue<string>("emailaddress1") : null;
				Workorder["hil_preferredtime"] = workOrderRequest.GetAttributeValue<OptionSetValue>("hil_preferreddaytime");
				Workorder["hil_preferreddate"] = workOrderRequest.GetAttributeValue<DateTime>("hil_preferreddate");
				Workorder["hil_address"] = new EntityReference("hil_address", addressId);
				Workorder["hil_productcategory"] = workOrderRequest.GetAttributeValue<EntityReference>("hil_productcategory");
				Workorder["hil_productcatsubcatmapping"] = workOrderRequest.GetAttributeValue<EntityReference>("hil_productsubcategory");
				Workorder["hil_consumertype"] = new EntityReference("hil_consumertype", new Guid("484897de-2abd-e911-a957-000d3af0677f"));
				Workorder["hil_consumercategory"] = new EntityReference("hil_consumercategory", new Guid("baa52f5e-2bbd-e911-a957-000d3af0677f"));
				Workorder["hil_natureofcomplaint"] = workOrderRequest.GetAttributeValue<EntityReference>("hil_natureofcomplaint");
				Workorder["hil_callsubtype"] = workOrderRequest.GetAttributeValue<EntityReference>("hil_callsubtype");
				Workorder["hil_quantity"] = workOrderRequest.GetAttributeValue<int>("hil_quantity");
				Workorder["hil_callertype"] = workOrderRequest.GetAttributeValue<OptionSetValue>("hil_callertype");
				//Workorder["hil_newserialnumber"] = workOrderRequest.GetAttributeValue<string>("hil_newserialnumber") +" | " + workOrderRequest.GetAttributeValue<string>("hil_dealercode");
				Workorder["hil_newserialnumber"] = workOrderRequest.GetAttributeValue<string>("hil_newserialnumber");
				Workorder["msdyn_serviceaccount"] = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43"));
				Workorder["msdyn_billingaccount"] = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43"));
				Workorder["hil_customercomplaintdescription"] = workOrderRequest.Contains("hil_customerdescriptionremarks") ? workOrderRequest.GetAttributeValue<string>("hil_customerdescriptionremarks") : null;
				Workorder["hil_emergencyremarks"] = workOrderRequest.Contains("hil_emergencyremarks") ? workOrderRequest.GetAttributeValue<string>("hil_emergencyremarks") : null;
				Workorder["hil_purchasedate"] = workOrderRequest.Contains("hil_purchasedate") ? workOrderRequest.GetAttributeValue<DateTime>("hil_purchasedate") : (DateTime?)null;
				Workorder["hil_emergencycall"] = workOrderRequest.Contains("hil_emergencycall") ? workOrderRequest.GetAttributeValue<OptionSetValue>("hil_emergencycall") : null;

                //Workorder["hil_addressdetails"] = workOrderRequest.Contains("hil_dealeremail") ? workOrderRequest.GetAttributeValue<string>("hil_dealeremail") : null;
                //Workorder["hil_b2bcompanyname"] = workOrderRequest.Contains("hil_b2bcompanyname") ? workOrderRequest.GetAttributeValue<string>("hil_b2bcompanyname") : null;

                Guid workorderid = service.Create(Workorder);

				#endregion
				#region WorkOrder_Request Updation

				Entity workOrder = service.Retrieve("msdyn_workorder", workorderid, new ColumnSet("msdyn_name"));
				if (workOrder.Id != null)
				{
					workOrderRequest["hil_workorder"] = workOrder.GetAttributeValue<string>("msdyn_name");
					workOrderRequest["hil_jobid"] = new EntityReference("msdyn_workorder", workorderid);
					workOrderRequest["hil_description"] = "Job created successfully.";
					service.Update(workOrderRequest);
				}
				//	tracingService.Trace("Work Order_ID :" + workOrder.GetAttributeValue<string>("msdyn_name"));
				//Console.WriteLine("Job Request ID : " + workOrderRequest.GetAttributeValue<EntityReference>("hil_jobid"));
				#endregion
			}		 
			else
			{
				workOrderRequest["hil_description"] = $"Duplicate Job Found :{ jobcheck[0].GetAttributeValue<string>("msdyn_name")}";
				//workOrderRequest["hil_jobid"] = new EntityReference("msdyn_workorder", jobcheck[0].Id);
				service.Update(workOrderRequest);
			}
		}

		private static void BulkJobUploader(IOrganizationService service, ITracingService tracingService, Entity workOrderRequest)
		{

			bool inputsValidated;
			string strInputsValidationsummary = string.Empty;

			try
			{
				#region JobRequest field Mapped in Excel

				#region Consumer field 
				string callingnumber = null;
				if (workOrderRequest.Contains("hil_name"))
				{ callingnumber = workOrderRequest.GetAttributeValue<string>("hil_name"); }

				string registerednumber = null;
				if (workOrderRequest.Contains("hil_registerednumber"))
				{ registerednumber = workOrderRequest.GetAttributeValue<string>("hil_registerednumber"); }

				string lastname = null;
				if (workOrderRequest.Contains("hil_lastname"))
				{ lastname = workOrderRequest.GetAttributeValue<string>("hil_lastname"); }

				string firstname = null;
				if (workOrderRequest.Contains("hil_firstname"))
				{ firstname = workOrderRequest.GetAttributeValue<string>("hil_firstname"); }

				string gender = null;
				if (workOrderRequest.Contains("hil_genderv1"))
				{ gender = workOrderRequest.GetAttributeValue<string>("hil_genderv1"); }

				string email = null;
				if (workOrderRequest.Contains("hil_email"))
				{ email = workOrderRequest.GetAttributeValue<string>("hil_email"); }

				string alternativenumber = null;
				if (workOrderRequest.Contains("hil_alternativenumber"))
				{ alternativenumber = workOrderRequest.GetAttributeValue<string>("hil_alternativenumber"); }

				#endregion

				#region Address Field

				string addressline1 = null;
				if (workOrderRequest.Contains("hil_addressline1"))
				{ addressline1 = workOrderRequest.GetAttributeValue<string>("hil_addressline1"); }

				string addressline2 = null;
				if (workOrderRequest.Contains("hil_addressline2"))
				{ addressline2 = workOrderRequest.GetAttributeValue<string>("hil_addressline2"); }

				string landmark = null;
				if (workOrderRequest.Contains("hil_landmark"))
				{ landmark = workOrderRequest.GetAttributeValue<string>("hil_landmark"); }

				string pincode = null;
				if (workOrderRequest.Contains("hil_pincodenamev1"))
				{ pincode = workOrderRequest.GetAttributeValue<string>("hil_pincodenamev1"); }

				#endregion

				#region Job Field

				string productsubcategory = null;
				if (workOrderRequest.Contains("hil_productsubcategorynamev1"))
				{ productsubcategory = workOrderRequest.GetAttributeValue<string>("hil_productsubcategorynamev1"); }

				string productcategory = null;
				if (workOrderRequest.Contains("hil_productcategorynamev1"))
				{ productcategory = workOrderRequest.GetAttributeValue<string>("hil_productcategorynamev1"); }

				string natureofcomplaint = null;
				if (workOrderRequest.Contains("hil_natureofcomplaintnamev1"))
				{ natureofcomplaint = workOrderRequest.GetAttributeValue<string>("hil_natureofcomplaintnamev1"); }

				string callsubtype = null;
				if (workOrderRequest.Contains("hil_callsubtypenamev1"))
				{ callsubtype = workOrderRequest.GetAttributeValue<string>("hil_callsubtypenamev1"); }

				int? quantity = null;
				if (workOrderRequest.Contains("hil_quantity"))
				{ quantity = workOrderRequest.GetAttributeValue<int>("hil_quantity"); }

				string calleretype = null;
				if (workOrderRequest.Contains("hil_callertypenamev1"))
				{ calleretype = workOrderRequest.GetAttributeValue<string>("hil_callertypenamev1"); }

				string preferreddaytime = null;
				if (workOrderRequest.Contains("hil_preferreddaytimenamev1"))
				{ preferreddaytime = workOrderRequest.GetAttributeValue<string>("hil_preferreddaytimenamev1"); }

				string dealercode = null;
				if (workOrderRequest.Contains("hil_dealercode"))
				{ dealercode = workOrderRequest.GetAttributeValue<string>("hil_dealercode"); }

				DateTime? purchasedate = null;
				if (workOrderRequest.Contains("hil_purchasedate"))
				{ purchasedate = workOrderRequest.GetAttributeValue<DateTime>("hil_purchasedate"); }

				DateTime? preferreddate = null;
				if (workOrderRequest.Contains("hil_preferreddate"))
				{ preferreddate = workOrderRequest.GetAttributeValue<DateTime>("hil_preferreddate"); }

				string emergencyname = null;
				if (workOrderRequest.Contains("hil_emergencycallnamev1"))
				{ emergencyname = workOrderRequest.GetAttributeValue<string>("hil_emergencycallnamev1"); }

				string emergencyremark = null;
				if (workOrderRequest.Contains("hil_emergencyremarks"))
				{ emergencyremark = workOrderRequest.GetAttributeValue<string>("hil_emergencyremarks"); }

				string customerdescriptionremarks = null;
				if (workOrderRequest.Contains("hil_customerdescriptionremarks"))
				{ customerdescriptionremarks = workOrderRequest.GetAttributeValue<string>("hil_customerdescriptionremarks"); }

				string newserialnumber = null;
				if (workOrderRequest.Contains("hil_newserialnumber"))
				{ newserialnumber = workOrderRequest.GetAttributeValue<string>("hil_newserialnumber"); }


				#endregion

				#endregion

				#region Validation JobRequestfield from Excel

				if (callingnumber == string.Empty || callingnumber == null) { strInputsValidationsummary += "\n callingnumber is required."; inputsValidated = false; }
				if (registerednumber == string.Empty || registerednumber == null) { strInputsValidationsummary += "\n registerednumber is required."; inputsValidated = false; }
				if (firstname == string.Empty || firstname == null) { strInputsValidationsummary += "\n firstname is required."; inputsValidated = false; }
				//	if (lastname == string.Empty || lastname == null) { strInputsValidationsummary += "\n lastname is required."; inputsValidated = false; }
			/*	if (gender == string.Empty || gender == null)
				{
					strInputsValidationsummary += "\n gender is required.";
					inputsValidated = false;
				}

				else
				{
					OptionSetValue optionSetValue = null;

					switch (gender.ToLower())
					{
						case "male": optionSetValue = new OptionSetValue(1); break;
						case "female": optionSetValue = new OptionSetValue(2); break;
						case "other": optionSetValue = new OptionSetValue(3); break;

					}
					workOrderRequest["hil_gender"] = optionSetValue;

					inputsValidated = true;
				}
			*/

				//	if (email == string.Empty || email == null) { strInputsValidationsummary += "\n email is required."; inputsValidated = false; }
				//	if (alternativenumber == string.Empty || alternativenumber == null) { strInputsValidationsummary += "\n alternativenumber is required."; inputsValidated = false; }
				if (addressline1 == string.Empty || addressline1 == null) { strInputsValidationsummary += "\n addressline1 is required."; inputsValidated = false; }
				//	if (addressline2 == string.Empty || addressline2 == null) { strInputsValidationsummary += "\n addressline2 is required."; inputsValidated = false; }
				//	if (landmark == string.Empty || landmark == null) { strInputsValidationsummary += "\n landmark is required."; inputsValidated = false; }
				if (pincode == string.Empty || pincode == null) { strInputsValidationsummary += "\n pincode is required."; inputsValidated = false; }

				else
				{
					EntityReference erPinCode = null;
					Entity entPincode = ExecuteScalar(service, "hil_pincode", "hil_name", pincode.Trim(), new string[] { "hil_pincodeid" });
					if (entPincode != null)
					{
						erPinCode = entPincode.ToEntityReference();
						var query = new QueryExpression("hil_businessmapping");
						query.TopCount = 1;
						query.ColumnSet.AddColumns("hil_businessmappingid", "hil_pincode");
						query.Criteria.AddCondition("hil_pincode", ConditionOperator.Equal, erPinCode.Id);
						query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);

						EntityCollection businessmapping = service.RetrieveMultiple(query);

						if (businessmapping.Entities.Count != 0)
						{
							foreach (var bm in businessmapping.Entities)
							{
								workOrderRequest["hil_pincode"] = new EntityReference("hil_businessmapping", bm.Id);
							}
							inputsValidated = true;
						}
					}
					else
					{
						strInputsValidationsummary += "/n Pincode Not Found";
						inputsValidated = false;
					}
				}

				if (productsubcategory == string.Empty || productsubcategory == null)
				{
					strInputsValidationsummary += "\n productsubcategory is required.";
					inputsValidated = false;
				}
				else
				{

					var query = new QueryExpression("hil_stagingdivisonmaterialgroupmapping");
					query.ColumnSet.AddColumns("hil_name", "hil_productcategorydivision", "hil_productsubcategorymg");
					query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, workOrderRequest.GetAttributeValue<string>("hil_productsubcategorynamev1").Trim());
					query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);

					EntityCollection productsubcategorys = service.RetrieveMultiple(query);
					if (productsubcategorys.Entities.Count != 0)
					{
						foreach (var psc in productsubcategorys.Entities)
						{
							workOrderRequest["hil_productsubcategory"] = new EntityReference("hil_stagingdivisonmaterialgroupmapping", psc.Id);
							workOrderRequest["hil_productcategory"] = new EntityReference("product", psc.GetAttributeValue<EntityReference>("hil_productcategorydivision").Id);
							Console.WriteLine("Prduct category & Sub Category" + psc.GetAttributeValue<EntityReference>("hil_productcategorydivision").Id + " " + psc.Id);
							productsubcategory = psc.GetAttributeValue<EntityReference>("hil_productsubcategorymg").Name;
						}
						//	service.Update(workOrderRequest);
						inputsValidated = true;
					}
					else
					{
						strInputsValidationsummary += "/n Product sub-category not found";
						inputsValidated = false;
						
					}
				}
				//	if (productcategory == string.Empty || productcategory == null) { strInputsValidationsummary += "\n productcategory is required."; inputsValidated = false; }
				if (natureofcomplaint == string.Empty || natureofcomplaint == null) { strInputsValidationsummary += "\n natureofcomplaint is required."; inputsValidated = false; }
				else
				{
					var query = new QueryExpression("hil_natureofcomplaint");

					// Add columns to query.ColumnSet
					query.ColumnSet.AddColumns("hil_callsubtype", "hil_name", "hil_natureofcomplaintid", "hil_relatedproduct");
					query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, natureofcomplaint.Trim());
					query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);

					var aa = query.AddLink("product", "hil_relatedproduct", "productid");
					aa.EntityAlias = "aa";
					aa.Columns.AddColumn("name");
					aa.LinkCriteria.AddCondition("name", ConditionOperator.Equal, productsubcategory);

					EntityCollection natureofcomplaints = service.RetrieveMultiple(query);

					if (natureofcomplaints.Entities.Count != 0)
					{
						foreach (var noc in natureofcomplaints.Entities)
						{
							workOrderRequest["hil_natureofcomplaint"] = new EntityReference("hil_natureofcomplaint", noc.Id);
							workOrderRequest["hil_callsubtype"] = new EntityReference("hil_callsubtype", noc.GetAttributeValue<EntityReference>("hil_callsubtype").Id);//noc.GetAttributeValue<EntityReference>("hil_callsubtype");
							Console.WriteLine("NOC & Subtype" + noc.GetAttributeValue<EntityReference>("hil_callsubtype").Id + " " + noc.Id);
						}

						inputsValidated = true;
					}
					else
					{
						strInputsValidationsummary += "/n Nature of complaint not found";
						inputsValidated = false;
							
					}



				}
				//	if (callsubtype == string.Empty || callsubtype == null) { strInputsValidationsummary += "\n callsubtype is required."; inputsValidated = false; }
				if (quantity == null) { strInputsValidationsummary += "\n quantity is required."; inputsValidated = false; }

				if (calleretype == string.Empty || calleretype == null)
				{
					strInputsValidationsummary += "\n calleretype is required.";
					inputsValidated = false;
				}

				else
				{
					OptionSetValue optionSetValue = null;

					switch (calleretype.ToLower())
					{
						case "email": optionSetValue = new OptionSetValue(910590005); break;
						case "customer": optionSetValue = new OptionSetValue(910590001); break;
						case "amc by cc": optionSetValue = new OptionSetValue(910590010); break;
						case "system": optionSetValue = new OptionSetValue(910590008); break;
						case "dealer": optionSetValue = new OptionSetValue(910590000); break;
						case "dealer helpdesk": optionSetValue = new OptionSetValue(910590004); break;
						case "havells employee": optionSetValue = new OptionSetValue(910590002); break;
						case "franchise": optionSetValue = new OptionSetValue(910590003); break;
						case "ecommerce": optionSetValue = new OptionSetValue(910590006); break;
						case "modern retail": optionSetValue = new OptionSetValue(910590007); break;
						case "brand store": optionSetValue = new OptionSetValue(910590009); break;
					}

					workOrderRequest["hil_callertype"] = optionSetValue;

					inputsValidated = true;
				}


				if (preferreddaytime == string.Empty || preferreddaytime == null)

				{
					strInputsValidationsummary += "\n preferreddaytime is required.";
					inputsValidated = false;
				}

				else
				{
					OptionSetValue optionSetValue = null;

					switch (preferreddaytime.ToLower())
					{
						case "morning": optionSetValue = new OptionSetValue(1); break;
						case "afternoon": optionSetValue = new OptionSetValue(2); break;
						case "evening": optionSetValue = new OptionSetValue(3); break;


					}

					workOrderRequest["hil_preferreddaytime"] = optionSetValue;

					inputsValidated = true;
				}

				//	if (dealercode == string.Empty || dealercode == null) { strInputsValidationsummary += "\n dealercode is required."; inputsValidated = false; }
				//	if (purchasedate == null) { strInputsValidationsummary += "\n purchasedate is required."; inputsValidated = false; }
				if (preferreddate == null) { strInputsValidationsummary += "\n preferreddate is required."; inputsValidated = false; }
				//	if (emergencyname == string.Empty || emergencyname == null) { strInputsValidationsummary += "\n emergencyname is required."; inputsValidated = false; }
				//	if (emergencyremark == string.Empty || emergencyremark == null) { strInputsValidationsummary += "\n emergencyremark is required."; inputsValidated = false; }
				//	if (customerdescriptionremarks == string.Empty || customerdescriptionremarks == null) { strInputsValidationsummary += "\n customerdescriptionremarks is required."; inputsValidated = false; }
				//	strInputsValidationsummary = newserialnumber ?? "newserialnumber is required";

				//	strInputsValidationsummary = string.IsNullOrWhiteSpace(newserialnumber) ? "newserialnumber is required" : "";

				/*	if (newserialnumber == string.Empty || newserialnumber == null)
					{
						strInputsValidationsummary += "\n newserialnumber is required.";
						//		inputsValidated = false; 
					//	Console.WriteLine("serial nmumber" + strInputsValidationsummary);
					}
				*/

				#endregion

				if (inputsValidated && strInputsValidationsummary == string.Empty)
				{
					Jobcretion(service, tracingService, workOrderRequest);
				}
				else
				{
					throw new InvalidPluginExecutionException(strInputsValidationsummary);
				}
			}
			catch (Exception ex) 
			{
				throw new InvalidPluginExecutionException(strInputsValidationsummary + ex.Message);
			}
		}

		private static Entity ExecuteScalar(IOrganizationService service, string entityName, string primaryField, string primaryFieldValue, string[] columns)
		{
			Entity retEntity = null;
			try
			{
				QueryExpression Query = new QueryExpression(entityName);
				Query.ColumnSet = new ColumnSet(columns);
				Query.Criteria = new FilterExpression(LogicalOperator.And);
				Query.Criteria.AddCondition(new ConditionExpression(primaryField, ConditionOperator.Equal, primaryFieldValue));
				Query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
				EntityCollection enCol = service.RetrieveMultiple(Query);
				if (enCol.Entities.Count >= 1)
				{
					retEntity = enCol.Entities[0];
				}
			}
			catch {}
			return retEntity;
		}
	}

}
