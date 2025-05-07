using System;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using System.Collections.Generic;
using System.Runtime.Serialization;
using System.Text.RegularExpressions;

namespace Havells_Plugin.TravelExpense
{
  public  class Common
    {
        public static void ShareWithBSH(Entity entity, IOrganizationService service)
        {
            try
            {
                using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                {
                    hil_travelexpense TravelExpense = entity.ToEntity<hil_travelexpense>();
                    if (TravelExpense.hil_submitted.Value == true)
                    {
                        TravelExpense = (hil_travelexpense)service.Retrieve(TravelExpense.LogicalName, TravelExpense.Id, new ColumnSet(true));
                        if (TravelExpense.hil_franchisee != null)
                        {
                            Guid fsBranchServiceHead = HelperAccount.GetBSHOwnerIdofAccount(TravelExpense.hil_franchisee.Id, service);
                            //Share the record with BSH
                            if (fsBranchServiceHead != Guid.Empty)
                            {
                                Helper.SharetoAccessTeam(new EntityReference(TravelExpense.LogicalName, TravelExpense.Id), fsBranchServiceHead, "Travel Expense Access Team Template", service);

                                //Update Travel expense status to Submitted
                                hil_travelexpense upTravelExpense = new hil_travelexpense();
                                upTravelExpense.hil_travelexpenseId = TravelExpense.hil_travelexpenseId.Value;
                                upTravelExpense.statuscode = new OptionSetValue(910590001);
                                service.Update(upTravelExpense);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.TravelExpense.Common.ShareWithBSH" + ex.Message);
            }

        }
    }
}
