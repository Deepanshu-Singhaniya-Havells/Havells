namespace Havells.D365.API.Contracts
{
    public static class ApiRoutes
    {
        public const string Base = "api";
        public static class WorkOrders
        {
            public const string GetWorkOrdersById =Base+ "/GetWorkOrdersById";
            public const string GetWorkOrdersByName = Base + "/GetWorkOrdersByName";
            public const string GetWorkOrderByCustRef = Base + "/GetWorkOrderByCustRef";
        }
        public static class Incident
        {
            public const string GetIncidentById = Base + "/GetIncidentById";
            public const string GetIndidentByWorkOrderId = Base + "/GetIndidentByWorkOrderId";
        }
        public static class WorkOrderProduct
        {
            public const string GetWOProductDetailByID = Base + "/GetWOProductDetailByID";
            public const string GetWOProductDetailByIncidentID = Base + "/GetWOProductDetailByIncidentID";
            public const string GetWOProductDetailByJobID = Base + "/GetWOProductDetailByJobID";
        }
        public static class WorkOrdService
        {
            public const string GetWOServiceDetailByJobID = Base + "/GetWOServiceDetailByJobID";
            public const string GetWOServiceDetailByID = Base + "/GetWOServiceDetailByID";
            public const string GetWOServiceDetailByIncidentID = Base + "/GetWOServiceDetailByIncidentID";
        }
        public static class CommonEntityApi
        {
            public const string GetD365ArchivedData = Base + "/GetD365ArchivedData";
        }
        public static class UserAuthentication
        {
            public const string AuthenticateUser = Base + "/Authenticateuser";
        }

    }
}
