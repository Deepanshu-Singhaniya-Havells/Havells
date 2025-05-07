using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsNewPlugin.Case
{
    public class CaseConstaints
    {
        public static readonly Guid _samparkDepartment = new Guid("7bf1705a-3764-ee11-8df0-6045bdaa91c3");
        public static readonly Guid _serviceDepartment = new Guid("ab3dbc3d-4e6e-ee11-8179-6045bdac526a");
        public static readonly Guid _ecommerceDepartment = new Guid("031d6bbb-8524-ef11-840a-000d3af09f9a");
        public static readonly int _callStaus_Answered = 9;

        public static readonly int _activityType_Assignment = 1;
        public static readonly int _activityType_Escalation = 2;
        public static readonly int _activityType_AssignmentSPOC = 4;

        public static readonly int _activityResolveBy_Assignee = 2;
        public static readonly int _activityResolveBy_SPOC = 1;
        //resolveby
    }
}
