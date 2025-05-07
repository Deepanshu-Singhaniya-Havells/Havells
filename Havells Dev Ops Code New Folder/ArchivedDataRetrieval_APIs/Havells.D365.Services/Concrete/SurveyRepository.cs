using Havells.D365.Data;
using Havells.D365.Entities.Common;
using Havells.D365.Entities.Survey.Request;
using Havells.D365.Entities.Survey.Response;
using Havells.D365.Services.Abstract;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.D365.Services.Concrete
{
    public class SurveyRepository : ISurveyRepository
    {
        private IConfiguration configuration;
        private readonly ILogger _logger;
        public SurveyRepository(IConfiguration configuration, ILoggerFactory logFactory)
        {
            this.configuration = configuration;
            this._logger = logFactory.CreateLogger<SurveyRepository>();
        }
        public SurveyDetail GetSurveyDetail(SurveyRequest request)
        {
            SurveyDetail objSurveyResponse = new SurveyDetail();
            var baseUrl = configuration["URLConnectionStrings:Survey"].ToString();
            _logger.LogInformation("baseUrl =" + baseUrl);
            string url = string.Concat(baseUrl, EndPoints.GetConsumerSurvey);
            _logger.LogInformation("url =" + url);
            _logger.LogInformation("requests "+JsonConvert.SerializeObject(request));
            IRestResponse response = RestServices.Post(url, JsonConvert.SerializeObject(request));
            _logger.LogInformation("api response status code " +JsonConvert.SerializeObject(response.StatusCode));
            _logger.LogInformation("response content from api "+ response.Content);
            if (((int)response.StatusCode)==200)
                objSurveyResponse = JsonConvert.DeserializeObject<SurveyDetail>(response.Content);
            _logger.LogInformation("response after deserilize "+JsonConvert.SerializeObject(objSurveyResponse));
            return objSurveyResponse;
        }

        public UpdateSurveyResponse SubmitSurvey(SurveyUpdateRequest request)
        {

            UpdateSurveyResponse objSurveyResponse = new UpdateSurveyResponse();
            var baseUrl = configuration["URLConnectionStrings:Survey"].ToString();
            string url = string.Concat(baseUrl, EndPoints.UpdateConsumerSurvey);
            IRestResponse response = RestServices.Post(url, JsonConvert.SerializeObject(request));
            if (((int)response.StatusCode) == 200)
                objSurveyResponse = JsonConvert.DeserializeObject<UpdateSurveyResponse>(response.Content);
            return objSurveyResponse;
        }
    }
}
