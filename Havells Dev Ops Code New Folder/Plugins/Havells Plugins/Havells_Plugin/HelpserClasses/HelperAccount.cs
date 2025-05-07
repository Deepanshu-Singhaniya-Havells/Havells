using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;

namespace Havells_Plugin
{
    public class HelperAccount
    {

        public static Guid GetBSHOwnerIdofAccount(Guid fsAccountId, IOrganizationService service)
        {
            using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
            {
                Guid fsResult = Guid.Empty;

                #region GetBSHAccount
                var obj = from _Account in orgContext.CreateQuery<Account>()
                          where _Account.Id == fsAccountId
                          select new
                          {
                              _Account.CustomerTypeCode,
                              _Account.Id,
                              _Account.OwnerId,
                              _Account.ParentAccountId,
                          };
                foreach (var iobj in obj)
                {
                    if (iobj.CustomerTypeCode != null && iobj.CustomerTypeCode.Value == ((int)Account_CustomerTypeCode.Branch))
                    {
                        return iobj.Id;
                    }
                    else
                    {
                        if (iobj.ParentAccountId != null)
                        {
                            fsResult = GetBSHOwnerIdofAccount(iobj.ParentAccountId.Id, service);
                        }
                    }
                }
                #endregion
                #region getBSHOwner
                if (fsResult != Guid.Empty)
                {
                    var obj1 = from _Account in orgContext.CreateQuery<Account>()
                               where _Account.Id == fsResult
                               select new
                               {
                                   _Account.Id,
                                   _Account.OwnerId,
                               };
                    foreach (var iobj1 in obj1)
                    {
                        fsResult = iobj1.OwnerId.Id;
                    }
                }
                #endregion

                return fsResult;
            }
        }

    }
}
