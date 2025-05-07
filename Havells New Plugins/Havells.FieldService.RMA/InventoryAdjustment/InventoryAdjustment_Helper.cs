using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.IdentityModel.Metadata;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.FieldService.RMA.InventoryAdjustment
{
    public class InventoryAdjustment_Helper
    {
        public enum InventoryAdjustmentStatus
        {
            Draft = 763620000,
            Submitted = 763620001,
            Approved = 763620002,
            Rejected = 763620003
        }
        public void UpdateQuantityOnInventoryAdjustmentProduct(IOrganizationService service, EntityReference parentRef)
        {
            try
            {
                QueryExpression query = new QueryExpression("hil_adjustmentproduct");
                query.ColumnSet = new ColumnSet("hil_adjustmentquantity", "hil_product", "hil_unit", "hil_inventoryadjustment", "ownerid", "hil_name", "hil_remarks");
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition("hil_inventoryadjustment", ConditionOperator.Equal, parentRef.Id);
                EntityCollection _entitys = service.RetrieveMultiple(query);
                ExecuteMultipleRequest _updateQuantityOnInventoryAdjustmentProduct = new ExecuteMultipleRequest()
                {
                    Settings = new ExecuteMultipleSettings()
                    {
                        ContinueOnError = false,
                        ReturnResponses = true
                    },
                    Requests = new OrganizationRequestCollection()
                };
                foreach (Entity _entity in _entitys.Entities)
                {
                    Entity entity = new Entity("msdyn_inventoryadjustmentproduct");

                    entity["msdyn_quantity"] = decimal.ToDouble((decimal)_entity["hil_adjustmentquantity"]);
                    entity["msdyn_product"] = _entity["hil_product"];
                    entity["msdyn_unit"] = _entity["hil_unit"];
                    entity["msdyn_inventoryadjustment"] = _entity["hil_inventoryadjustment"];
                    entity["ownerid"] = _entity["ownerid"];
                    if (_entity.Contains("hil_name"))
                        entity["msdyn_name"] = _entity["hil_name"];
                    if (_entity.Contains("hil_remarks"))
                        entity["hil_remarks"] = _entity["hil_remarks"];
                    CreateRequest createRequest = new CreateRequest { Target = entity };
                    _updateQuantityOnInventoryAdjustmentProduct.Requests.Add(createRequest);
                }
                ExecuteMultipleResponse responseWithResults = (ExecuteMultipleResponse)service.Execute(_updateQuantityOnInventoryAdjustmentProduct);
                foreach (var r in responseWithResults.Responses)
                {
                    if (r.Response != null)
                        throw new InvalidPluginExecutionException("Success");
                    else if (r.Fault != null)
                        throw new InvalidPluginExecutionException("Update Quantity Error " + r.Fault.Message);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}
