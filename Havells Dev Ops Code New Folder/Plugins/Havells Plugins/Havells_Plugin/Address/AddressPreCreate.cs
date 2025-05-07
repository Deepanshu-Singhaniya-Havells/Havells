using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Havells_Plugin.Address
{
    public class AddressPreCreate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == "hil_address" && (context.MessageName.ToUpper() == "CREATE"))
                {
                    Entity entity = (Entity)context.InputParameters["Target"];

                    if (!entity.Contains("hil_street1")) {
                        throw new InvalidPluginExecutionException("Address Line 1 is required.");
                    }
                    if (!entity.Contains("hil_businessgeo"))
                    {
                        throw new InvalidPluginExecutionException("Business Geo Mapping is required.");
                    }
                    else {
                        Guid _busnessGeoMapping = entity.GetAttributeValue<EntityReference>("hil_businessgeo").Id;
                        string _fetchxml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='hil_businessmapping'>
                            <attribute name='hil_businessmappingid' />
                            <filter type='and'>
                                <condition attribute='hil_businessmappingid' operator='eq' value='{_busnessGeoMapping}' />
                                <condition attribute='hil_branch' operator='not-null' />
                                <condition attribute='hil_pincode' operator='not-null' />
                                <condition attribute='statecode' operator='eq' value='0' />
                            </filter>
                            </entity>
                            </fetch>";

                        EntityCollection _entCol = service.RetrieveMultiple(new FetchExpression(_fetchxml));
                        if (_entCol.Entities.Count == 0) {
                            throw new InvalidPluginExecutionException("Business Geo Mapping is Inactive/Not Properly defined.");
                        }
                    }
                    //hil_address iPermanent = entity.ToEntity<hil_address>();
                    //if (iPermanent.hil_AddressType != null)
                    //{
                    //    string PermAddress = string.Empty;
                    //    if (iPermanent.hil_Street1 != null)
                    //        PermAddress = iPermanent.hil_Street1.ToUpper() + ", ";
                    //    if (iPermanent.hil_Street2 != null)
                    //        PermAddress = PermAddress + iPermanent.hil_Street2.ToUpper() + ", ";
                    //    if (iPermanent.hil_Street3 != null)
                    //        PermAddress = PermAddress + iPermanent.hil_Street3.ToUpper() + ", ";
                    //    if (iPermanent.hil_CIty != null)
                    //        PermAddress = PermAddress + iPermanent.hil_CIty.Name.ToUpper() + ", ";
                    //    if (iPermanent.hil_State != null)
                    //        PermAddress = PermAddress + iPermanent.hil_State.Name.ToUpper();
                    //    if (iPermanent.hil_PinCode != null)
                    //        PermAddress = PermAddress + " - " + iPermanent.hil_PinCode.Name;
                    //    iPermanent.hil_FullAddress = PermAddress;
                    //    if (iPermanent.hil_AddressType.Value == 1)
                    //        iPermanent.hil_name = "Permanent";
                    //    else if(iPermanent.hil_AddressType.Value == 2)
                    //        iPermanent.hil_name = "Alternate";
                    //    if (iPermanent.hil_Customer != null && iPermanent.hil_AddressType.Value == 1)
                    //    {
                    //        Contact _enCont = (Contact)service.Retrieve(Contact.EntityLogicalName, iPermanent.hil_Customer.Id, new ColumnSet("hil_permanentaddress"));
                    //        _enCont.hil_PermanentAddress = PermAddress;
                    //        service.Update(_enCont);
                    //    }
                    //}
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}
