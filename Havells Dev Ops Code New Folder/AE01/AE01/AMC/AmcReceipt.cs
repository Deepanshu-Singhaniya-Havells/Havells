
using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace AE01.AMC
{

    public class AmcReceipt
    {
        private IOrganizationService service;
        public AmcReceipt(IOrganizationService _service)
        {
            this.service = _service;
        }
        internal void Test()
        {
            Entity salesOrderLine = service.Retrieve("salesorderdetail", new Guid("1771306a-1b74-ef11-ac20-7c1e5205cf66"), new ColumnSet(true));
            StampDiscount(salesOrderLine);
        }

        private EntityCollection GetDiscountPercentage(EntityReference productCategory, EntityReference productSubCategory, string currentFormatedDate, string ageing)
        {
            string fetchDiscount = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                                        <entity name='hil_amcdiscountmatrix'>
                                                        <attribute name='hil_amcdiscountmatrixid' />
                                                        <attribute name='hil_discounttype' />
                                                        <attribute name='hil_discper' />
                                                        <order attribute='hil_name' descending='false' />
                                                        <filter type='and'>
                                                        <condition attribute='hil_appliedto' operator='eq' value='{{03B5A2D6-CC64-ED11-9562-6045BDAC526A}}' />
                                                        <condition attribute='hil_productcategory' operator='eq' value='{productCategory.Id}' />
                                                        <filter type='or'>
                                                        <condition attribute='hil_productsubcategory' operator='null' />
                                                        <condition attribute='hil_productsubcategory' operator='eq' value='{productSubCategory.Id}' />
                                                        </filter>
                                                        <condition attribute='hil_validfrom' operator='on-or-before' value='{currentFormatedDate}' />
                                                        <condition attribute='hil_validto' operator='on-or-after' value='{currentFormatedDate}' />
                                                        <condition attribute='hil_productaegingstart' operator='le' value='{ageing}' />
                                                        <condition attribute='hil_productageingend' operator='ge' value='{ageing}' />
                                                        <condition attribute='statecode' operator='eq' value='0' />
                                                        <condition attribute='hil_discper' operator='gt' value='0' />
                                                        </filter>
                                                        </entity>
                                                        </fetch>";

            return service.RetrieveMultiple(new FetchExpression(fetchDiscount));
        }

        public void StampDiscount(Entity orderItem)
        {
            if (orderItem.Contains("salesorderid"))
            {
                Entity salesOrder = service.Retrieve("salesorder", orderItem.GetAttributeValue<EntityReference>("salesorderid").Id, new ColumnSet(true));

                string orderType = salesOrder.GetAttributeValue<EntityReference>("hil_ordertype").Name;

                if (orderType.Contains("AMC"))
                {
                    Entity product = service.Retrieve("product", orderItem.GetAttributeValue<EntityReference>("productid").Id, new ColumnSet("hil_division", "hil_materialgroup"));
                    EntityReference productCategory = product.Contains("hil_division") ? product.GetAttributeValue<EntityReference>("hil_division") : null;
                    EntityReference productSubCategory = product.Contains("hil_materialgroup") ? product.GetAttributeValue<EntityReference>("hil_materialgroup") : null;
                    DateTime orderItemCreationDate = orderItem.Contains("createdon") ? orderItem.GetAttributeValue<DateTime>("createdon") : DateTime.MinValue;
                    string currentFormatedDate = DateTime.Now.ToString("yyyy-MM-dd");

                    if (orderItem.Contains("hil_customerasset"))
                    {
                        Entity customerAsset = service.Retrieve("msdyn_customerasset", orderItem.GetAttributeValue<EntityReference>("hil_customerasset").Id, new ColumnSet("hil_invoicedate"));

                        DateTime invoiceDate = customerAsset.Contains("hil_invoicedate") ? customerAsset.GetAttributeValue<DateTime>("hil_invoicedate") : DateTime.MinValue;

                        if (productCategory != null && productSubCategory != null && orderItemCreationDate != DateTime.MinValue)
                        {
                            int ageing = (orderItemCreationDate - invoiceDate).Days;

                            EntityCollection discountCollection = GetDiscountPercentage(productCategory, productSubCategory, currentFormatedDate, ageing.ToString());

                            if (discountCollection.Entities.Count > 0)
                            {
                                decimal discountPercentage = discountCollection[0].Contains("hil_discper") ? discountCollection[0].GetAttributeValue<decimal>("hil_discper") : 0;
                                decimal amount = orderItem.GetAttributeValue<Money>("baseamount").Value;
                                orderItem["hil_eligiblediscount"] = new Money(amount * discountPercentage / 100);
                            }
                            //service.Update(orderItem); 
                        }

                    }
                    else if (orderItem.Contains("hil_assetserialnumber"))
                    {
                        DateTime invoiceDate = orderItem.Contains("hil_invoicedate") ? orderItem.GetAttributeValue<DateTime>("hil_invoicedate") : DateTime.MinValue;

                        if (productCategory != null && productSubCategory != null && invoiceDate != DateTime.MinValue && orderItemCreationDate != DateTime.MinValue)
                        {
                            int ageing = (orderItemCreationDate - invoiceDate).Days;

                            EntityCollection discountCollection = GetDiscountPercentage(productCategory, productSubCategory, currentFormatedDate, ageing.ToString());
                            if (discountCollection.Entities.Count > 0)
                            {
                                decimal discountPercentage = discountCollection[0].Contains("hil_discper") ? discountCollection[0].GetAttributeValue<decimal>("hil_discper") : 0;

                                decimal amount = orderItem.GetAttributeValue<Money>("baseamount").Value;
                                orderItem["hil_eligiblediscount"] = new Money(amount * discountPercentage / 100);
                                //service.Update(orderItem); 
                            }

                        }
                    }

                }

            }
        }
    }

}
