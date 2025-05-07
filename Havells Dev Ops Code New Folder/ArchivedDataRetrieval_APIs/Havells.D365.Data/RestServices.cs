using RestSharp;

namespace Havells.D365.Data
{
    public sealed class RestServices
    {
        
        public static IRestResponse Post(string url,string requestBody)
        {
            var client = new RestClient(url);
            client.Timeout = -1;
            var request = new RestRequest(Method.POST);
            request.AddHeader("Content-Type", "application/json");
            request.AddParameter("application/json", requestBody, ParameterType.RequestBody);
            var response= client.Execute(request);
            return response;
        }

    }
}
