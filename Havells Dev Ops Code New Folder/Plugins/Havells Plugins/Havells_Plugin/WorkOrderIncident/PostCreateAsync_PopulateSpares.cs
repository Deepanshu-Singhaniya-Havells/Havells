using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;

namespace Havells_Plugin.WorkOrderIncident
{
    public class PostCreateAsync_PopulateSpares : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            try
            {
                if (context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == msdyn_workorderincident.EntityLogicalName
                    && context.MessageName.ToUpper() == "CREATE" && context.Depth <= 1)
                {

                    Entity entity = (Entity)context.InputParameters["Target"];
                    msdyn_workorderincident _entWrkInc = entity.ToEntity<msdyn_workorderincident>();

                    //tracingService.Trace($@"Step 1: Retrieving Model Num bind with Customer Asset. {DateTime.Now.ToString()}");
                    msdyn_customerasset CustomerAssest = (msdyn_customerasset)service.Retrieve(msdyn_customerasset.EntityLogicalName, entity.GetAttributeValue<EntityReference>("msdyn_customerasset").Id, new ColumnSet("msdyn_product"));
                    //tracingService.Trace($@"Step 2: Retrieval done. {DateTime.Now.ToString()}");
                    
                    if (CustomerAssest != null && CustomerAssest.msdyn_Product != null)
                    {
                        Guid Model = CustomerAssest.msdyn_Product.Id;
                        tracingService.Trace($@"Step 3: Fetching Starts - Incident Products based on Incident Type. {DateTime.Now.ToString()}");

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
                                      <condition attribute='name' operator='eq' value='{Common._restrictSpareRoleName}' />
                                    </filter>
                                  </link-entity>
                                </link-entity>
                              </entity>
                            </fetch>";
                            EntityCollection _entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                            if (_entCol.Entities.Count == 0)
                            {
                                IfCauseProductsPresent(service, _entWrkInc, Model, tracingService);
                            }
                        }
                        else {
                            IfCauseProductsPresent(service, _entWrkInc, Model, tracingService);
                        }
                        tracingService.Trace($@"Step 4: Fetching Completed - Incident Products based on Incident Type. {DateTime.Now.ToString()}");
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.WorkOrderIncident.PostCreateAsync_PopulateSpares: " + ex.Message + ": " + DateTime.Now.ToString());
            }
        }
        #region If Cause Products Present
        public static void IfCauseProductsPresent(IOrganizationService service, msdyn_workorderincident _WoInc, Guid Model, ITracingService tracingService)
        {
            //#region Added by Kuldeep Khare 03/May/2020 to fetch Job's Call Subtype 
            //bool _isAMCJob = false;
            //if (_WoInc.msdyn_WorkOrder != null)
            //{
            //    Entity entObj = service.Retrieve(msdyn_workorder.EntityLogicalName, _WoInc.msdyn_WorkOrder.Id, new ColumnSet("hil_callsubtype"));
            //    if (entObj.GetAttributeValue<EntityReference>("hil_callsubtype").Id == new Guid("55A71A52-3C0B-E911-A94E-000D3AF06CD4"))
            //    {
            //        _isAMCJob = true;
            //    }
            //}
            //#endregion
            tracingService.Trace($@"Step 3.1: Fetching Starts - Incident Products based on Incident Type. {DateTime.Now.ToString()}");
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
                tracingService.Trace($@"Step 3.2: Looping Through All Incident Products and creating Work Order Product Lines. {DateTime.Now.ToString()}");
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
                tracingService.Trace($@"Step 3.2: Execute Multiple Starts to create Work Order Product Lines. {DateTime.Now.ToString()}");
                ExecuteMultipleResponse responseWithResults = (ExecuteMultipleResponse)service.Execute(requestWithResults);
                tracingService.Trace($@"Step 3.3: Execute Multiple Ends to create Work Order Product Lines. {DateTime.Now.ToString()}");
            }
            else
            {
                tracingService.Trace($@"Step 5.1: Fetching Starts - Quering Service BOM Table based on Model Num. {DateTime.Now.ToString()}");
                PopulateWorkOrderProduct(service, _WoInc, Model, tracingService);
                tracingService.Trace($@"Step 5.2: Fetching Ends - Quering Service BOM Table based on Model Num. {DateTime.Now.ToString()}");
            }
        }
        public static void PopulateWorkOrderProduct(IOrganizationService service, msdyn_workorderincident _WoInc, Guid Model, ITracingService tracingService)
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
