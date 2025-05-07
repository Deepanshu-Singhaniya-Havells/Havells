using Havells.CRM.DynamicsPOSyncToSAP;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.Linq;

namespace DynamicsPOSyncToSAP
{
    public class SyncSAPtoD365BillingDetail
    {
        private readonly ServiceClient _service;
        public SyncSAPtoD365BillingDetail(ServiceClient service)
        {
            _service = service;
        }
        public void UpdateAMCStagingTable()
        {
            try
            {
                IntegrationConfiguration integrationConfiguration = HelperClass.GetIntegrationConfiguration(_service, "AMC Billing Details");
                string FromDate = GetLastRunDate();// new DateTime(2024, 11, 15).ToString("yyyy-MM-dd"); 
                string ToDate = DateTime.Now.ToString("yyyy-MM-dd");

                ParamRequest objParam = new ParamRequest
                {
                    FROM_DATE = FromDate,
                    TO_DATE = ToDate,
                    IM_FLAG = "R",
                    CallNumber = "",
                    SerialNumber = ""
                };

                var resp = JsonConvert.DeserializeObject<AmcBillingDetailList>(HelperClass.CallAPI(integrationConfiguration, JsonConvert.SerializeObject(objParam), "POST"));
                if (resp?.LT_TABLE != null && resp.LT_TABLE.Count > 0)
                {
                    AmcBillingDetails objAmcBilling = resp.LT_TABLE.Where(m => m.SERIAL == "67HDG60401191").FirstOrDefault();
                    //foreach (AmcBillingDetails objAmcBilling in abc)
                    //{
                        EntityReference _amcPlan = GetAmcPlan(objAmcBilling.MATNR);
                        string _fetchXML = GetFetchXml(objAmcBilling, _amcPlan);

                        EntityCollection objEntity = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (objEntity.Entities.Count == 0)
                        {
                            CreateInvoiceDetails(objAmcBilling, _amcPlan);
                            Console.WriteLine("SAP Bill Number:{0} added Successfully.", objAmcBilling.VBELN_B);
                        }
                        else
                        {
                            Console.WriteLine("SAP Bill Number:{0} already exist.", objAmcBilling.VBELN_B);
                        }
                   // }
                }
                else
                {
                    Console.WriteLine("No record found from SAP.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error " + ex.Message);
            }
        }

        private EntityReference GetAmcPlan(string matnr)
        {
            QueryExpression qsCType = new QueryExpression("product")
            {
                ColumnSet = new ColumnSet("name"),
                NoLock = true
            };
            qsCType.Criteria.AddCondition("name", ConditionOperator.Equal, matnr);
            EntityCollection AmcPlanGuid = _service.RetrieveMultiple(qsCType);
            return AmcPlanGuid.Entities.Count > 0 ? AmcPlanGuid.Entities[0].ToEntityReference() : null;
        }

        private string GetFetchXml(AmcBillingDetails objAmcBilling, EntityReference _amcPlan)
        {
            if (_amcPlan != null)
            {
                return $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
							  <entity name='hil_amcstaging'>
								<attribute name='hil_amcstagingid' />
								<filter type='and'>
								  <condition attribute='hil_name' operator='eq' value='{objAmcBilling.VBELN_B}' />
								  <condition attribute='hil_serailnumber' operator='eq' value='{objAmcBilling.SERIAL}' />
								  <filter type='or'>
									<condition attribute='hil_amcplan' operator='eq' value='{_amcPlan.Id}' />
									<condition attribute='hil_amcplannameslt' operator='eq' value='{objAmcBilling.MATNR}' />
								  </filter>
								</filter>
							  </entity>
							</fetch>";
            }
            else
            {
                return $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
							  <entity name='hil_amcstaging'>
								<attribute name='hil_amcstagingid' />
								<filter type='and'>
								  <condition attribute='hil_name' operator='eq' value='{objAmcBilling.VBELN_B}' />
								  <condition attribute='hil_serailnumber' operator='eq' value='{objAmcBilling.SERIAL}' />
								  <condition attribute='hil_amcplannameslt' operator='eq' value='{objAmcBilling.MATNR}' />
								</filter>
							  </entity>
							</fetch>";
            }
        }

        private void CreateInvoiceDetails(AmcBillingDetails objAmcBilling, EntityReference _amcPlan)
        {
            if (objAmcBilling.SERIAL == "67HDG60401191")
            {
                Entity invoiceDetails = new Entity("hil_amcstaging")
                {
                    ["hil_serailnumber"] = objAmcBilling.SERIAL,
                    ["hil_callid"] = objAmcBilling.CALLNO,
                    ["hil_mobilenumber"] = objAmcBilling.MOBILENO,
                    ["hil_name"] = objAmcBilling.VBELN_B,
                    ["hil_amcplannameslt"] = objAmcBilling.MATNR,
                    ["hil_warrantystartdate"] = Convert.ToDateTime(objAmcBilling.AMCSTART),
                    ["hil_warrantyenddate"] = Convert.ToDateTime(objAmcBilling.AMCEND),
                    ["hil_sapbillingdate"] = Convert.ToDateTime(objAmcBilling.ERDAT),
                    ["hil_amcpurchasedate"] = Convert.ToDateTime(objAmcBilling.CUSTPDATE),
                    //                   ["hil_isdeleteflag"] = objAmcBilling.DEL_FLAG == "x",
                    ["hil_branchmemorandumcode"] = objAmcBilling.KUNNR,
                    ["hil_salesordernumber"] = objAmcBilling.VBELN_S,
                    ["hil_amcplan"] = _amcPlan
                };
                _service.Create(invoiceDetails);
            }
        }

        private string GetLastRunDate()
        {
            string _lastAMCPurchaseDate = "2024-04-28";
            try
            {
                QueryExpression Query = new QueryExpression("hil_amcstaging")
                {
                    ColumnSet = new ColumnSet("modifiedon", "createdon"),
                    Criteria = new FilterExpression(LogicalOperator.And)
                };
                Query.Criteria.AddCondition("modifiedon", ConditionOperator.NotNull);
                Query.TopCount = 1;
                Query.AddOrder("modifiedon", OrderType.Descending);

                EntityCollection Found = _service.RetrieveMultiple(Query);
                if (Found.Entities.Count > 0)
                {
                    DateTime _date = Found.Entities[0].GetAttributeValue<DateTime>("modifiedon").AddMinutes(330);
                    _lastAMCPurchaseDate = _date.ToString("yyyy-MM-dd");
                }
                return _lastAMCPurchaseDate;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error " + ex.Message);
            }
            return _lastAMCPurchaseDate;
        }
    }
}
