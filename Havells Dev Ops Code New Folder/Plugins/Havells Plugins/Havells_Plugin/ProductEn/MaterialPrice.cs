using System;
using System.Collections.Generic;

namespace Havells_Plugin
{
    public class MaterialPrice
    {
        public string Condition { get; set; }
        public string FromDate { get; set; }
        public int IsInitialLoad { get; set; }
        public string MaterialCode { get; set; }
        public string ToDate { get; set; }
    }

    public class MaterialPriceResponseObj
    {
        public string MATNR { get; set; }
        public string KSCHL { get; set; }
        public double KBETR { get; set; }
        public string KONWA { get; set; }
        public string VKORG { get; set; }
        public DateTime DATAB { get; set; }
        public DateTime DATBI { get; set; }
        public string DELETE_FLAG { get; set; }
        public string CTIMESTAMP { get; set; }
        public string CreatedBy { get; set; }
        public string MTIMESTAMP { get; set; }
        public string ModifyBy { get; set; }
    }

    public class MaterialPriceResponseRoot
    {
        public object Result { get; set; }
        public List<MaterialPriceResponseObj> Results { get; set; }
        public bool Success { get; set; }
        public object Message { get; set; }
        public object ErrorCode { get; set; }
    }
}
