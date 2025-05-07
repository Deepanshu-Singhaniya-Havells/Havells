using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.CRM.UnitWarrantyRefresh
{
    internal class Enums_Warranty
    {
    }
    public enum WarrentyType
    {
        Standard = 1,
        Extended = 2,
        AMC = 3,
        Comprehensive = 4,
        Component = 5,
        Service = 6,
        SpecialScheme = 7
    }
    public enum WarrentySubStatus
    {
        Standard = 1,
        Extended = 2,
        SpecialScheme = 3,
        UnderAMC = 4
    }
    public enum WarrentyStatus
    {
        InWarranty = 1,
        OutWarranty = 2,
        WarrantyVoid = 3,
        NAforWarranty = 4
    }
    public enum AssetStatusCode
    {
        PendingforApproval = 910590000,
        ProductApproved = 910590001
    }
}
