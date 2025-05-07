using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class SFA_DivisionCallType
    {
        [DataMember]
        public string DivisionName { get; set; }
        [DataMember]
        public string DivisionCode { get; set; }
        [DataMember]
        public string CallTypeCode { get; set; }
        [DataMember]
        public string CallTypeName { get; set; }

        public List<SFA_DivisionCallType> GetDivisionCallTypeSetup()
        {
            List<SFA_DivisionCallType> lstDivisionCallType = new List<SFA_DivisionCallType>();
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    lstDivisionCallType.Add(new SFA_DivisionCallType() {CallTypeCode = "I",CallTypeName = "Installation",DivisionCode = "31",DivisionName = "CRABTREE AUTOMATION"});
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
    }
}
