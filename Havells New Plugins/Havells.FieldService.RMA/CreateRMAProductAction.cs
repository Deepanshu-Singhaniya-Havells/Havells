using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Services;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace Havells.FieldService.RMA
{
    public class CreateRMAProductAction : IPlugin
    {
        public static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                string JobProductIDs = context.InputParameters["JobProductIDs"].ToString();
                string RMAID = context.InputParameters["RMAID"].ToString();

                Entity RMAEntity = service.Retrieve("msdyn_rma", new Guid(RMAID), new ColumnSet("ownerid", "msdyn_serviceaccount"));
                EntityReference rmaOwner = RMAEntity.GetAttributeValue<EntityReference>("ownerid");
                Entity account = service.Retrieve(RMAEntity.GetAttributeValue<EntityReference>("msdyn_serviceaccount").LogicalName,
                     RMAEntity.GetAttributeValue<EntityReference>("msdyn_serviceaccount").Id, new ColumnSet("defaultpricelevelid"));
                EntityReference priceList = null;
                if (account.Contains("defaultpricelevelid"))
                    priceList = account.GetAttributeValue<EntityReference>("defaultpricelevelid");
                else
                    throw new InvalidPluginExecutionException("Price List is not mapped.");
                EntityReference wareHouse = GetWarehouse(rmaOwner, service);

                string[] jobproductIdsArray = JobProductIDs.Split(',');
                string value = "";
                foreach (string jobproductId in jobproductIdsArray)
                {
                    value = value + "<value>{" + jobproductId.Trim() + "}</value>";
                }
                string fetchXML = $@"<fetch version=""1.0"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"">
                                      <entity name=""msdyn_workorderproduct"">
                                        <attribute name=""msdyn_workorderproductid"" />
                                        <attribute name=""hil_replacedpartdescription"" />
                                        <attribute name=""hil_replacedpart"" />
                                        <attribute name=""msdyn_customerasset"" />
                                        <attribute name=""msdyn_quantity"" />
                                        <attribute name=""hil_partamount"" />
                                        <order attribute=""hil_replacedpart"" descending=""false"" />
                                        <filter type=""and"">
                                          <condition attribute=""msdyn_workorderproductid"" operator=""in"">
                                            {value}
                                          </condition>
                                        </filter>
                                        <link-entity name=""product"" from=""productid"" to=""hil_replacedpart"" visible=""false"" link-type=""outer"" alias=""product"">
                                          <attribute name=""description"" />
                                          <attribute name=""defaultuomid"" />
                                        </link-entity>
                                      </entity>
                                    </fetch>";
                EntityCollection entityCollection = service.RetrieveMultiple(new FetchExpression(fetchXML));
                foreach (Entity entity in entityCollection.Entities)
                {
                    Entity RMAProduct = new Entity("msdyn_rmaproduct");
                    RMAProduct["msdyn_rma"] = RMAEntity.ToEntityReference();
                    RMAProduct["msdyn_quantitytoreturn"] = entity.GetAttributeValue<double>("msdyn_quantity");
                    RMAProduct["msdyn_woproduct"] = entity.ToEntityReference();
                    RMAProduct["msdyn_product"] = entity.GetAttributeValue<EntityReference>("hil_replacedpart");
                    RMAProduct["msdyn_customerasset"] = entity.GetAttributeValue<EntityReference>("msdyn_customerasset");
                    RMAProduct["msdyn_pricelist"] = priceList;
                    RMAProduct["msdyn_processingaction"] = new OptionSetValue(690970001);
                    RMAProduct["ownerid"] = rmaOwner;
                    if (entity.Contains("product.description"))
                        RMAProduct["msdyn_description"] = entity.GetAttributeValue<AliasedValue>("product.description").Value.ToString();
                    if (entity.Contains("product.defaultuomid"))
                        RMAProduct["msdyn_unit"] = (EntityReference)entity.GetAttributeValue<AliasedValue>("product.defaultuomid").Value;
                    RMAProduct["msdyn_returntowarehouse"] = wareHouse;
                    RMAProduct["msdyn_unitamount"] = new Money(entity.GetAttributeValue<Decimal>("hil_partamount"));
                    service.Create(RMAProduct);
                    service.Update(
                        new Entity(entity.LogicalName, entity.Id)
                        {
                            Attributes = new AttributeCollection() {
                                new KeyValuePair<string, object>("hil_returned", new OptionSetValue())
                            }
                        }
                    );

                }
                context.OutputParameters["Status"] = "Sucess !";
                context.OutputParameters["Message"] = "RMA lines created";
            }
            catch (Exception ex)
            {
                context.OutputParameters["Status"] = "Error !";
                context.OutputParameters["Message"] = "Error " + ex.Message;
            }
        }
        private EntityReference GetWarehouse(EntityReference rmaOwner, IOrganizationService service)
        {
            EntityReference wareHouse = null;
            QueryExpression Query = new QueryExpression("bookableresource");
            Query.ColumnSet = new ColumnSet("msdyn_warehouse");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("userid", ConditionOperator.Equal, rmaOwner.Id);
            Query.AddOrder("createdon", OrderType.Descending);
            EntityCollection bookableColl = service.RetrieveMultiple(Query);
            if (bookableColl.Entities.Count != 1)
            {
                throw new InvalidPluginExecutionException("Bookable resource is not mapped.");
            }
            else
            {
                if (bookableColl[0].Contains("msdyn_warehouse"))
                {
                    wareHouse = bookableColl[0].GetAttributeValue<EntityReference>("msdyn_warehouse");
                }
                else
                {
                    throw new InvalidPluginExecutionException("Warehouse is not mapped.");
                }
            }
            return wareHouse;
        }

    }
}
