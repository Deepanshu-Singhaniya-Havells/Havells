using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.Customer_Asset
{
    public class SFAGetDivisionCallTypeSetup : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion
            try
            {
                string jsonResponse = string.Empty;
                List<SFA_DivisionCallType> lstDivisionCallType = new List<SFA_DivisionCallType>();
                jsonResponse = JsonSerializer.Serialize(GetDivisionCallTypeSetup(service));
                context.OutputParameters["data"] = jsonResponse;
                return;

            }
            catch (Exception ex)
            {
                List<SFA_DivisionCallType> lstDivisionCallType = new List<SFA_DivisionCallType>();
                {
                    new SFA_DivisionCallType().DivisionCode = "Error";
                    new SFA_DivisionCallType().DivisionName = "D365 Internal Server Error: " + ex.Message;
                };
                context.OutputParameters["data"] = JsonSerializer.Serialize(GetDivisionCallTypeSetup(service));
            }
        }
        public List<SFA_DivisionCallType> GetDivisionCallTypeSetup(IOrganizationService service)
        {
            List<SFA_DivisionCallType> lstDivisionCallType = new List<SFA_DivisionCallType>();
            try
            {
                if (service != null)
                {
                    lstDivisionCallType.Add(new SFA_DivisionCallType() { CallTypeCode = "I", CallTypeName = "Installation", DivisionCode = "31", DivisionName = "CRABTREE AUTOMATION" });
                    lstDivisionCallType.Add(new SFA_DivisionCallType() { CallTypeCode = "I", CallTypeName = "Installation", DivisionCode = "30", DivisionName = "CRABTREE EWA" });
                    lstDivisionCallType.Add(new SFA_DivisionCallType() { CallTypeCode = "D", CallTypeName = "Demo", DivisionCode = "49", DivisionName = "HAVELLS AIR COOLER" });
                    lstDivisionCallType.Add(new SFA_DivisionCallType() { CallTypeCode = "I", CallTypeName = "Installation", DivisionCode = "44", DivisionName = "HAVELLS AQUA" });
                    lstDivisionCallType.Add(new SFA_DivisionCallType() { CallTypeCode = "I", CallTypeName = "Installation", DivisionCode = "45", DivisionName = "HAVELLS FAN" });
                    lstDivisionCallType.Add(new SFA_DivisionCallType() { CallTypeCode = "D", CallTypeName = "Demo", DivisionCode = "47", DivisionName = "HAVELLS SDA" });
                    lstDivisionCallType.Add(new SFA_DivisionCallType() { CallTypeCode = "B", CallTypeName = "Both", DivisionCode = "48", DivisionName = "HAVELLS WATER HEATER" });
                    lstDivisionCallType.Add(new SFA_DivisionCallType() { CallTypeCode = "B", CallTypeName = "Both", DivisionCode = "81", DivisionName = "LLOYD AIR CONDITIONER" });
                    lstDivisionCallType.Add(new SFA_DivisionCallType() { CallTypeCode = "B", CallTypeName = "Both", DivisionCode = "82", DivisionName = "LLOYD LED TELEVISION" });
                    lstDivisionCallType.Add(new SFA_DivisionCallType() { CallTypeCode = "I", CallTypeName = "Installation", DivisionCode = "84", DivisionName = "LLOYD WASHING MACHIN" });
                    lstDivisionCallType.Add(new SFA_DivisionCallType() { CallTypeCode = "I", CallTypeName = "Installation", DivisionCode = "67", DivisionName = "SOLAR LSP" });
                    lstDivisionCallType.Add(new SFA_DivisionCallType() { CallTypeCode = "I", CallTypeName = "Installation", DivisionCode = "96", DivisionName = "STANDARD FAN" });
                    lstDivisionCallType.Add(new SFA_DivisionCallType() { CallTypeCode = "I", CallTypeName = "Installation", DivisionCode = "97", DivisionName = "STANDARD WATERHEATER" });
                    return lstDivisionCallType;
                }
                else
                {
                    lstDivisionCallType.Add(new SFA_DivisionCallType() { DivisionCode = "ERROR", DivisionName = "D365 Service Unavailable" });
                    return lstDivisionCallType;
                }
            }
            catch (Exception ex)
            {
                lstDivisionCallType.Add(new SFA_DivisionCallType() { DivisionCode = "ERROR", DivisionName = "D365 Internal Server Error : " + ex.Message });
                return lstDivisionCallType;
            }
        }
        public class SFA_DivisionCallType
        {
            public string DivisionName { get; set; }
            public string DivisionCode { get; set; }
            public string CallTypeCode { get; set; }
            public string CallTypeName { get; set; }

        }
    }
}
