using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Havells_Plugin.Address
{
    public class AddressPostUpdate : IPlugin
    {
        private static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            // Obtain the execution context from the service provider.
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == hil_address.EntityLogicalName && (context.MessageName.ToUpper() == "UPDATE" && context.Depth == 1))
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    hil_address _permAdd = (hil_address)service.Retrieve(hil_address.EntityLogicalName, entity.Id, new ColumnSet(true)); //entity.ToEntity<hil_address>();
                    
                    if (_permAdd.hil_AddressType != null)
                    {
                        string PermAddress = string.Empty;
                        if (_permAdd.hil_Street1 != null)
                            PermAddress = _permAdd.hil_Street1.ToUpper() + ", ";
                        if (_permAdd.hil_Street2 != null)
                            PermAddress = PermAddress + _permAdd.hil_Street2.ToUpper() + ", ";
                        if (_permAdd.hil_Street3 != null)
                            PermAddress = PermAddress + _permAdd.hil_Street3.ToUpper() + ", ";
                        if (_permAdd.hil_CIty != null)
                            PermAddress = PermAddress + _permAdd.hil_CIty.Name.ToUpper() + ", ";
                        if (_permAdd.hil_State != null)
                            PermAddress = PermAddress + _permAdd.hil_State.Name.ToUpper();
                        if (_permAdd.hil_PinCode != null)
                            PermAddress = PermAddress + " - " + _permAdd.hil_PinCode.Name;
                        hil_address _iperm = new hil_address();
                        _iperm.Id = _permAdd.Id;
                        _iperm.hil_FullAddress = PermAddress;
                        if (_permAdd.hil_AddressType.Value == 1)
                            _iperm.hil_name = "Permanent";
                        else if (_permAdd.hil_AddressType.Value == 2)
                            _iperm.hil_name = "Alternate";
                        service.Update(_iperm);

                        if (_permAdd.hil_Customer != null && _permAdd.hil_AddressType.Value == 1)
                        {
                            Contact _enCont = (Contact)service.Retrieve(Contact.EntityLogicalName, _permAdd.hil_Customer.Id, new ColumnSet("hil_permanentaddress"));
                            _enCont.hil_PermanentAddress = PermAddress;
                            service.Update(_enCont);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}
