using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;

namespace D365WebJobs
{
    public class PostCreateAsync_PopulateSpares
    {
        public void Execute(IOrganizationService service)
        {
            try
            {
                string _restrictSpareRoleName = "Restrict Spare Part Prepopulating";
                Entity entity = service.Retrieve("msdyn_workorderincident", new Guid("f3dfc9be-51c6-ee11-9079-000d3a3e404e"), new ColumnSet(true));
                msdyn_workorderincident _entWrkInc = entity.ToEntity<msdyn_workorderincident>();
                msdyn_customerasset CustomerAssest = (msdyn_customerasset)service.Retrieve(msdyn_customerasset.EntityLogicalName, entity.GetAttributeValue<EntityReference>("msdyn_customerasset").Id, new ColumnSet("msdyn_product"));
                if (CustomerAssest != null && CustomerAssest.msdyn_Product != null)
                {
                    Guid Model = CustomerAssest.msdyn_Product.Id;

                    if (_entWrkInc.Contains("ownerid"))
                    {
                        string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                              <entity name='systemuser'>
                                <attribute name='systemuserid' />
                                <filter type='and'>
                                  <condition attribute='systemuserid' operator='eq' value='{_entWrkInc.GetAttributeValue<EntityReference>("ownerid").Id}' />
                                </filter>
                                <link-entity name='systemuserroles' from='systemuserid' to='systemuserid' visible='false' intersect='true'>
                                  <link-entity name='role' from='roleid' to='roleid' alias='ae'>
                                    <filter type='and'>
                                      <condition attribute='name' operator='eq' value='{_restrictSpareRoleName}' />
                                    </filter>
                                  </link-entity>
                                </link-entity>
                              </entity>
                            </fetch>";
                        EntityCollection _entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (_entCol.Entities.Count == 0)
                        {
                            IfCauseProductsPresent(service, _entWrkInc, Model);
                        }
                    }
                    else
                    {
                        IfCauseProductsPresent(service, _entWrkInc, Model);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        #region If Cause Products Present
        public static void IfCauseProductsPresent(IOrganizationService service, msdyn_workorderincident _WoInc, Guid Model)
        {
            OptionSetValue iChargable = new OptionSetValue();
            QueryExpression Query = new QueryExpression()
            {
                EntityName = msdyn_incidenttypeproduct.EntityLogicalName,
                ColumnSet = new ColumnSet("msdyn_product", "msdyn_quantity", "msdyn_lineorder", "hil_chargeableornot", "hil_isserialized"),
                Criteria = {
                    Conditions = {
                        new ConditionExpression("msdyn_incidenttype", ConditionOperator.Equal, _WoInc.msdyn_IncidentType.Id),
                        new ConditionExpression("hil_model", ConditionOperator.Equal, Model)
                    }
                },
                NoLock = true
            };

            EntityCollection PossibleFaultProductColl = service.RetrieveMultiple(Query);
            if (PossibleFaultProductColl.Entities != null && PossibleFaultProductColl.Entities.Count > 0)
            {
                ExecuteMultipleRequest requestWithResults = new ExecuteMultipleRequest()
                {
                    // Assign settings that define execution behavior: continue on error, return responses. 
                    Settings = new ExecuteMultipleSettings()
                    {
                        ContinueOnError = false,
                        ReturnResponses = true
                    },
                    // Create an empty organization request collection.
                    Requests = new OrganizationRequestCollection()
                };
                foreach (msdyn_incidenttypeproduct pCause in PossibleFaultProductColl.Entities)
                {
                    msdyn_workorderproduct WoPdt = new msdyn_workorderproduct();
                    WoPdt.msdyn_CustomerAsset = _WoInc.msdyn_CustomerAsset;
                    WoPdt.msdyn_WorkOrderIncident = new EntityReference(msdyn_workorderincident.EntityLogicalName, _WoInc.msdyn_workorderincidentId.Value);
                    WoPdt.msdyn_WorkOrder = _WoInc.msdyn_WorkOrder != null ? _WoInc.msdyn_WorkOrder : null;
                    WoPdt.msdyn_Product = pCause.msdyn_Product != null ? pCause.msdyn_Product : null;
                    WoPdt.hil_MaxQuantity = pCause.msdyn_Quantity != null ? Convert.ToDecimal(pCause.msdyn_Quantity) : 1;
                    WoPdt["hil_priority"] = pCause.msdyn_LineOrder != null ? Convert.ToString(pCause.msdyn_LineOrder) : string.Empty;
                    if (pCause.Contains("hil_chargeableornot"))
                    {
                        iChargable = pCause.GetAttributeValue<OptionSetValue>("hil_chargeableornot");
                        WoPdt["hil_chargeableornot"] = (OptionSetValue)pCause["hil_chargeableornot"];
                        if (iChargable.Value == 1)
                        {
                            WoPdt.hil_WarrantyStatus = new OptionSetValue(2);
                        }
                    }
                    WoPdt.hil_Quantity = 1;
                    WoPdt.msdyn_Quantity = 1;

                    Product Pdt = (Product)service.Retrieve(Product.EntityLogicalName, pCause.msdyn_Product.Id, new ColumnSet("name", "description", "hil_amount"));

                    string Uq = Pdt.Description != null ? (pCause.msdyn_Product.Name + "-" + Pdt.Description) : pCause.msdyn_Product.Name + "-";
                    WoPdt["hil_part"] = Uq;

                    if (Pdt.hil_Amount != null)
                    {
                        WoPdt.msdyn_TotalAmount = Pdt.hil_Amount;
                        WoPdt.hil_PartAmount = Pdt.hil_Amount.Value;
                    }
                    if (pCause.Contains("hil_isserialized"))
                    {
                        WoPdt.hil_IsSerialized = (OptionSetValue)pCause["hil_isserialized"];
                    }
                    CreateRequest createRequest = new CreateRequest() { Target = WoPdt };
                    requestWithResults.Requests.Add(createRequest);
                }
                ExecuteMultipleResponse responseWithResults = (ExecuteMultipleResponse)service.Execute(requestWithResults);
            }
            else
            {
                PopulateWorkOrderProduct(service, _WoInc, Model);
            }
        }
        public static void PopulateWorkOrderProduct(IOrganizationService service, msdyn_workorderincident _WoInc, Guid Model)
        {
            try
            {
                QueryExpression Query = new QueryExpression()
                {
                    EntityName = hil_servicebom.EntityLogicalName,
                    ColumnSet = new ColumnSet("hil_isserialized", "hil_quantity", "hil_product", "hil_priority", "hil_chargeableornot"),
                    Criteria = {
                    Conditions = {
                        new ConditionExpression("hil_productcategory", ConditionOperator.Equal, Model),
                        new ConditionExpression("statecode", ConditionOperator.Equal, 0)
                        }
                    },
                    NoLock = true
                };
                EntityCollection ServiceBomColl = service.RetrieveMultiple(Query);
                if (ServiceBomColl.Entities != null && ServiceBomColl.Entities.Count >= 1)
                {
                    ExecuteMultipleRequest requestWithResults = new ExecuteMultipleRequest()
                    {
                        // Assign settings that define execution behavior: continue on error, return responses. 
                        Settings = new ExecuteMultipleSettings()
                        {
                            ContinueOnError = false,
                            ReturnResponses = true
                        },
                        // Create an empty organization request collection.
                        Requests = new OrganizationRequestCollection()
                    };

                    for (int i = 0; i < ServiceBomColl.Entities.Count; i++)
                    {
                        hil_servicebom Srv = (hil_servicebom)ServiceBomColl.Entities[i];
                        msdyn_workorderproduct WoPdt = new msdyn_workorderproduct();
                        WoPdt.msdyn_CustomerAsset = _WoInc.msdyn_CustomerAsset;
                        WoPdt.msdyn_Product = Srv.hil_Product != null ? Srv.hil_Product : throw new InvalidPluginExecutionException("Part null in Serice bom");

                        Product Pdt1 = (Product)service.Retrieve(Product.EntityLogicalName, Srv.hil_Product.Id, new ColumnSet("name", "description", "hil_amount"));

                        string Uq = Pdt1.Description != null ? Srv.hil_Product.Name + "-" + Pdt1.Description : Srv.hil_Product.Name + "-";
                        WoPdt["hil_part"] = Uq;
                        WoPdt["hil_priority"] = Srv.Contains("hil_priority") ? (string)Srv["hil_priority"] : string.Empty;
                        WoPdt.msdyn_WorkOrderIncident = new EntityReference(msdyn_workorderincident.EntityLogicalName, _WoInc.Id);
                        WoPdt.msdyn_WorkOrder = _WoInc.msdyn_WorkOrder;
                        WoPdt.hil_MaxQuantity = Srv.hil_quantity != null ? Srv.hil_quantity : 1;
                        WoPdt.msdyn_Quantity = 1;
                        if(Srv.Contains("hil_chargeableornot") && Srv.hil_chargeableornot != null)
                        {
                            WoPdt["hil_chargeableornot"] = Srv.hil_chargeableornot;
                            if(Srv.hil_chargeableornot.Value == 1)
                            {
                                WoPdt.hil_WarrantyStatus = new OptionSetValue(2);
                            }
                        }
                        if (Pdt1.hil_Amount != null)
                        {
                            WoPdt.msdyn_TotalAmount = Pdt1.hil_Amount;
                            WoPdt["hil_partamount"] = Pdt1.hil_Amount.Value;
                        }
                        if (Srv.Contains("hil_isserialized"))
                        {
                            WoPdt["hil_isserialized"] = (OptionSetValue)Srv["hil_isserialized"];
                        }
                        CreateRequest createRequest = new CreateRequest() { Target = WoPdt };
                        requestWithResults.Requests.Add(createRequest);
                        //service.Create(WoPdt);
                    }
                    ExecuteMultipleResponse responseWithResults = (ExecuteMultipleResponse)service.Execute(requestWithResults);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.WorkOrderIncident.PostCreate.PopulateWorkOrderProduct" + ex.Message);
            }
        }
        #endregion
    }
}
