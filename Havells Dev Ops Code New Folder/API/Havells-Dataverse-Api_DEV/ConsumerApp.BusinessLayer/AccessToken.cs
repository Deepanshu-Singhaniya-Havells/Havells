using Newtonsoft.Json;
using RestSharp;
using System;
using System.Runtime.Serialization;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class AccessToken 
    {
        public string getAccessToken()
        {
            var client = new RestClient("https://login.microsoftonline.com/7b7dc2f5-4e6a-4004-96dd-6c7923625b25/oauth2/token");
            client.Timeout = -1;
            var request = new RestRequest(Method.POST);
             request.AlwaysMultipartFormData = true;
            request.AddParameter("client_id", "41623af4-f2a7-400a-ad3a-ae87462ae44e");
            request.AddParameter("client_secret", "r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=");
            request.AddParameter("resource", "https://havells.crm8.dynamics.com");
            request.AddParameter("grant_type", "client_credentials");
            IRestResponse response = client.Execute(request);
            Console.WriteLine(response.Content);
            Root myDeserializedClass = JsonConvert.DeserializeObject<Root>(response.Content);
            return myDeserializedClass.access_token;
        }
    }
    [DataContract]
    public class Root
    {
        [DataMember]
        public string token_type { get; set; }
        [DataMember] 
        public string expires_in { get; set; }
        [DataMember] 
        public string ext_expires_in { get; set; }
        [DataMember] 
        public string expires_on { get; set; }
        [DataMember] 
        public string not_before { get; set; }
        [DataMember] 
        public string resource { get; set; }
        [DataMember] 
        public string access_token { get; set; }
    }
}
