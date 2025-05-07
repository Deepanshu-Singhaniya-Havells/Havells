using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Crm.Sdk.Messages;
using Havells_Plugin.Consumer_App_Bridge;

namespace Havells_Plugin
{
    public class HelperWorkOrder
    {
        //    public static Validate CreateWorkOrder(IOrganizationService service, Guid CustAsst, Guid Customer, Guid ServAcct, Guid PrimaryIcdntTyp)
        //    {
        //        Havells_Plugin.Consumer_App_Bridge.Validate VD = new Validate();
        //        msdyn_workorder _WrkOd = new msdyn_workorder();
        //        _WrkOd.msdyn_CustomerAsset = new EntityReference(msdyn_customerasset.EntityLogicalName, CustAsst);
        //        _WrkOd.hil_CustomerRef = new EntityReference(Contact.EntityLogicalName, Customer);
        //        _WrkOd.msdyn_ServiceAccount = new EntityReference(Account.EntityLogicalName, ServAcct);
        //        _WrkOd.msdyn_PrimaryIncidentType = new EntityReference(msdyn_incidenttype.EntityLogicalName, PrimaryIcdntTyp);
        //        Guid _WoID = service.Create(_WrkOd);
        //        if (_WoID != new Guid("00000000-0000-0000-0000-000000000000"))
        //        {
        //            VD.StatusCode = true;
        //            VD.WOUniqueID = _WoID.ToString();
        //        }
        //        else
        //        {
        //            VD.StatusCode = false;
        //            VD.ExceptionCode = "10";
        //            VD.ExceptionDesc = "CRM Error";
        //        }
        //        return (VD);
        //    }
        //
    }
}
