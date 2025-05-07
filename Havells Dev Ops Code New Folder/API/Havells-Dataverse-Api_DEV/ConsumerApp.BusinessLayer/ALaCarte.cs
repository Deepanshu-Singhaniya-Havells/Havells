using System;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System.Runtime.Serialization;
using Microsoft.Xrm.Tooling.Connector;


namespace ConsumerApp.BusinessLayer
{
	[DataContract]
	public class ALaCarte
	{
		public OrderCreateresponse CreateSalesOrder(CreateSalesOrderModel salesOrder)
		{
			OrderCreateresponse objOrderCreateresponse = new OrderCreateresponse();
			try
			{
				IOrganizationService service = ConnectToCRM.GetOrgService();
				if (service != null)
				{
					if (salesOrder != null)
					{

						if (string.IsNullOrWhiteSpace(salesOrder.OrderType))
						{
							objOrderCreateresponse.APIStatus = "OrderType is required";
							return objOrderCreateresponse;
						}

						if (string.IsNullOrEmpty(salesOrder.ConsumerID))
						{
							var fetch = $@"<fetch>
                                  <entity name=""contact"">
                                    <attribute name=""fullname"" /> 
                                        <attribute name=""mobilephone"" />
                                    <filter>
                                      <condition attribute=""mobilephone"" operator=""eq"" value=""{salesOrder.ConsumerMobile}"" />
                                    </filter>
                                  </entity>
                                </fetch>";

							EntityCollection contactcoll = service.RetrieveMultiple(new FetchExpression(fetch));

							if (contactcoll.Entities.Count == 0)
							{
								Entity entConsumer = new Entity("contact");

								if (!string.IsNullOrEmpty(salesOrder.ConsumerMobile))
								{
									entConsumer["mobilephone"] = salesOrder.ConsumerMobile;
								}
								else
								{
									objOrderCreateresponse.APIStatus = "ConsumerMobile is required";
									return objOrderCreateresponse;
								}

								if (!string.IsNullOrEmpty(salesOrder.ConsumerName))
								{
									entConsumer["firstname"] = salesOrder.ConsumerName;
								}
								else
								{
									objOrderCreateresponse.APIStatus = "ConsumerName is required";
									return objOrderCreateresponse;
								}

								if (!string.IsNullOrEmpty(salesOrder.ConsumerEmail))
								{
									entConsumer["emailaddress1"] = salesOrder.ConsumerEmail;
								}
								else
								{
									objOrderCreateresponse.APIStatus = "ConsumerEmail is required";
									return objOrderCreateresponse;
								}

								entConsumer["hil_consumersource"] = new OptionSetValue(1);
								entConsumer["hil_subscribeformessagingservice"] = true;
								entConsumer["hil_preferredlanguageforcommunication"] = new EntityReference("hil_preferredlanguageforcommunication", new Guid("d825675f-37db-ec11-a7b5-6045bdad294c"));

								Guid contactId = service.Create(entConsumer);
								salesOrder.ConsumerID = Convert.ToString(contactId);
							}
							else
							{
								salesOrder.ConsumerID = contactcoll[0].Id.ToString();
							}
						}

						Entity SOEntity = new Entity("salesorder");
						SOEntity["customerid"] = new EntityReference("contact", new Guid(salesOrder.ConsumerID));
						SOEntity["msdyn_psastatusreason"] = new OptionSetValue(192350000);
						SOEntity["transactioncurrencyid"] = new EntityReference("pricelevel", new Guid("68a6a9ca-6beb-e811-a96c-000d3af05828"));
						SOEntity["ownerid"] = new EntityReference("systemuser", new Guid(salesOrder.Technicianid));
						SOEntity["msdyn_ordertype"] = new OptionSetValue(690970002);
						SOEntity["msdyn_account"] = new EntityReference("account", new Guid("d166ba69-65da-ec11-a7b5-6045bdad2a19"));
						SOEntity["hil_source"] = new OptionSetValue(3);
						#region 02092024 Address validation
						if (!string.IsNullOrWhiteSpace(salesOrder.addressid))
						{
							Entity EntAdd = service.Retrieve("hil_address", new Guid(salesOrder.addressid), new ColumnSet(true));
							if (EntAdd != null)
							{
								SOEntity["hil_serviceaddress"] = new EntityReference("hil_address", EntAdd.Id);
							}
							else
							{
								objOrderCreateresponse.APIStatus = "serivce address not found";
							}
						}
						if (!string.IsNullOrWhiteSpace(salesOrder.Pincode) && string.IsNullOrWhiteSpace(salesOrder.addressid))
						{
							var query = new QueryExpression("hil_businessmapping");
							query.TopCount = 1;
							query.ColumnSet.AddColumns("hil_businessmappingid", "hil_name", "createdon");
							query.Criteria.AddCondition("hil_businessmappingid", ConditionOperator.NotNull);
							query.AddOrder("hil_name", OrderType.Ascending);
							var ac = query.AddLink("hil_pincode", "hil_pincode", "hil_pincodeid");
							ac.EntityAlias = "ac";
							ac.LinkCriteria.AddCondition("hil_name", ConditionOperator.Equal, salesOrder.Pincode);

							EntityCollection businessmapping = service.RetrieveMultiple(query);
							if (businessmapping.Entities.Count > 0)
							{
								Entity addresspin = new Entity("hil_address");
								addresspin["hil_customer"] = new EntityReference("contact", new Guid(salesOrder.ConsumerID));
								addresspin["hil_addresstype"] = new OptionSetValue(1);
								addresspin["hil_street1"] = salesOrder.addressline1;
								addresspin["hil_businessgeo"] = new EntityReference("hil_businessmapping", businessmapping.Entities[0].Id);
								addresspin["hil_name"] = salesOrder.addressline1;
								SOEntity["hil_serviceaddress"] = new EntityReference("hil_address", service.Create(addresspin));

							}
							else
							{
								objOrderCreateresponse.APIStatus = "Invalid Pincode";
								return objOrderCreateresponse;
							}
						}
						#endregion

						#region OrderType validation

						QueryExpression ordertypequery = new QueryExpression("hil_ordertype");
						ordertypequery.ColumnSet.AddColumns("hil_ordertype", "hil_ordertypeid", "hil_pricelist");
						ordertypequery.Criteria.AddCondition("hil_ordertypeid", ConditionOperator.Equal, salesOrder.OrderType.ToString());

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
							objOrderCreateresponse.APIStatus = "Order Type not found";
							return objOrderCreateresponse;
						}
						#endregion

						Guid orderID = service.Create(SOEntity);
						if (orderID != null)
						{
							Entity socon = service.Retrieve("contact", SOEntity.GetAttributeValue<EntityReference>("customerid").Id, new ColumnSet("firstname", "mobilephone", "emailaddress1"));
							string _contactmobilephone = socon.Contains("mobilephone") ? socon.GetAttributeValue<string>("mobilephone") : string.Empty;
							string _contactfirstname = socon.Contains("firstname") ? socon.GetAttributeValue<string>("firstname") : string.Empty;
							string _contactemailaddress = socon.Contains("emailaddress1") ? socon.GetAttributeValue<string>("emailaddress1") : string.Empty;

							ColumnSet columnSalesOrder = new ColumnSet("ordernumber", "createdon", "msdyn_psastatusreason", "name");
							Entity entitySalesOrder = service.Retrieve("salesorder", orderID, columnSalesOrder);
							string _orderNumber = entitySalesOrder.GetAttributeValue<string>("name");   //OrderNumber;
							string _orderDate = entitySalesOrder.GetAttributeValue<DateTime>("createdon").ToString();
							string _orderStatus = entitySalesOrder.GetAttributeValue<OptionSetValue>("msdyn_psastatusreason").Value.ToString();

							foreach (SalesOrderProductDetailsModel prod in salesOrder.ProductDetails)
							{
								Entity soitem = new Entity("salesorderdetail");
								soitem["salesorderid"] = new EntityReference("salesorder", orderID);
								soitem["productid"] = new EntityReference("product", new Guid(prod.ProductID));
								soitem["quantity"] = Convert.ToDecimal(prod.Quantity);
								soitem["priceperunit"] = new Money(Convert.ToDecimal(prod.PricePerUnit));
								soitem["baseamount"] = new Money(Convert.ToDecimal(prod.Amount));
								soitem["uomid"] = new EntityReference("uom", new Guid("0359d51b-d7cf-43b1-87f6-fc13a2c1dec8"));
								soitem["ownerid"] = new EntityReference("systemuser", new Guid(salesOrder.Technicianid));
								Guid orderLineID = service.Create(soitem);
							}
							objOrderCreateresponse.OrderNumer = _orderNumber;
							objOrderCreateresponse.OrderDate = _orderDate;
							objOrderCreateresponse.ConsumerMobile = _contactmobilephone; //salesOrder.ConsumerMobile;
							objOrderCreateresponse.ConsumerName = _contactfirstname;//salesOrder.ConsumerName;
							objOrderCreateresponse.OrderStatus = _orderStatus;
							objOrderCreateresponse.APIStatus = "Success";
							objOrderCreateresponse.ProductDetails = salesOrder.ProductDetails;
						}
						else
						{
							objOrderCreateresponse.APIStatus = "Failed to create Order.";

						}
					}
				}
				else
				{
					return new OrderCreateresponse { APIStatus = "D365 Service Unavailable." };
				}
				return objOrderCreateresponse;
			}
			catch (Exception ex)
			{
				return new OrderCreateresponse { APIStatus = ex.Message };
			}
		}
	}
	[DataContract]
	public class CreateSalesOrderModel
	{
		[DataMember]
		public string OrderType { get; set; }
		[DataMember]
		public string ConsumerID { get; set; }
		[DataMember]
		public string ConsumerMobile { get; set; }
		[DataMember]
		public string ConsumerSalutation { get; set; }
		[DataMember]
		public string ConsumerName { get; set; }
		[DataMember]
		public string ConsumerEmail { get; set; }
		[DataMember]
		public string SubTotal { get; set; }
		[DataMember]
		public string Discount { get; set; }
		[DataMember]
		public string NetAmount { get; set; }
		[DataMember]
		public string Technicianid { get; set; }
		[DataMember]
		public string Pincode { get; set; }
		[DataMember]
		public string addressline1 { get; set; }
		[DataMember]
		public string addressid { get; set; }
		[DataMember]
		public List<SalesOrderProductDetailsModel> ProductDetails { get; set; }
	}
	[DataContract]
	public class SalesOrderProductDetailsModel
	{
		//[DataMember]
		//public string OrderNumer { get; set; }
		[DataMember]
		public string ProductID { get; set; }
		[DataMember]
		public string Quantity { get; set; }
		[DataMember]
		public string PricePerUnit { get; set; }
		[DataMember]
		public string Amount { get; set; }
	}
	[DataContract]
	public class OrderCreateresponse
	{
		[DataMember]
		public string ConsumerMobile { get; set; }
		[DataMember]
		public string ConsumerName { get; set; }
		[DataMember]
		public string OrderNumer { get; set; }
		[DataMember]
		public string OrderDate { get; set; }
		[DataMember]
		public string OrderStatus { get; set; }
		[DataMember]
		public List<SalesOrderProductDetailsModel> ProductDetails { get; set; }
		[DataMember]
		public string APIStatus { get; set; }

	}
	[DataContract]
	public class SendAlaCartSMSRequest
	{
		[DataMember]
		public string OrderId { get; set; }
		[DataMember]
		public string mobile { get; set; }
		[DataMember]
		public string Amount { get; set; }
	}

}