using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class Ecom_Priceinfo
    {
        public Ecom_PriceDetailsResponse GetEcom_Priceinfo(Ecom_PriceDetailsRequest request)
        {
            
            return new Ecom_PriceDetailsResponse()
            {
                ModelNumber = "Success"
            };
            //Ecom_PriceDetails obj = new Ecom_PriceDetails();
            //float amount = await GetAmount(request.ModelNumber);
            //if (amount == -1F)
            //{
            //    obj.Message = "Not Able to call the API";
            //    obj.Amount = -1F;
            //}
            //else if (amount == -2)
            //{
            //    obj.Message = "Model Number does not exist";
            //    obj.Amount = -2F;
            //}
            //else
            //{
            //    obj.Message = "Success";
            //    obj.Amount = amount;
            //}
            //return obj;
        }
        private async Task<float> GetAmount(string modelNumber)
        {
            float ans = -1F;
            string amount = "";
            var client = new HttpClient();
            var request = new HttpRequestMessage(HttpMethod.Get, "https://middlewaredev.havells.com:50001/RESTAdapter/ecom_priceinfo?IM_FLAG=&IM_PROJECT=D365");
            request.Headers.Add("Authorization", "Basic RDM2NV9IQVZFTExTOkRFVkQzNjVAMTIzNA==");
            // request.Headers.Add("Cookie", "JSESSIONID=S-8sbup2i2P7m-XQfu34jiWrea1PjgGenHkA_SAPS48sCbzvJ7oRjokDRLt8Pe02; saplb_*=(J2EE7969920)7969950");
            var content = new StringContent("{    \r\n     \"LT_TABLE\" : {\r\n         \"MATNR\" :  " + modelNumber + "\r\n    }\r\n}", null, "application/json");
            request.Content = content;
            var response = await client.SendAsync(request);

            if (response.IsSuccessStatusCode)
            {
                Ecom_PriceinfoResponse obj = JsonConvert.DeserializeObject<Ecom_PriceinfoResponse>(await response.Content.ReadAsStringAsync());
                if (obj != null)
                {
                    foreach (var item in obj.LT_TABLE)
                    {
                        if (item.KSCHL != null && item.KSCHL == "ZWEB")
                        {
                            amount = item.KBETR;
                        }
                    }
                }
                else
                {
                    ans = -2F;
                }
            }
            if (!string.IsNullOrEmpty(amount)) ans = float.Parse(amount);
            return ans;
        }
    }

    [DataContract]
    public class Ecom_PriceDetailsResponse
    {
        [DataMember]
        public string Message { get; set; }
        [DataMember]
        public float Amount { get; set; }
        [DataMember]
        public string ModelNumber { get; set; }

    }
    [DataContract]
    public class Ecom_PriceDetailsRequest
    {
        [DataMember]
        public string ModelNumber { get; set; }
    }
    public class LTTABLE
    {
        public string MATNR { get; set; }
        public string KSCHL { get; set; }
        public string DATAB { get; set; }
        public string KBETR { get; set; }
        public string KONWA { get; set; }
        public string DATBI { get; set; }
        public string DELETE_FLAG { get; set; }
        public string CREATEDBY { get; set; }
        public string CTIMESTAMP { get; set; }
        public string MODIFYBY { get; set; }
        public string MTIMESTAMP { get; set; }
    }
    public class Ecom_PriceinfoResponse
    {
        public List<LTTABLE> LT_TABLE { get; set; }
    }
}
