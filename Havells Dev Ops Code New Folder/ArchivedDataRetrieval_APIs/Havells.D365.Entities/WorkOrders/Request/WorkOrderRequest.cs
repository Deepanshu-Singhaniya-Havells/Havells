namespace Havells.D365.Entities.WorkOrders.Request
{
    public class WorkOrderRequest
    {
        public string CustomerName { get; set; }
        public string CustomerReferenceNo { get; set; }
        public string WorkOrderId { get; set; }
        public int PageNo { get; set; }
        public int Size { get; set; }
    }
}
