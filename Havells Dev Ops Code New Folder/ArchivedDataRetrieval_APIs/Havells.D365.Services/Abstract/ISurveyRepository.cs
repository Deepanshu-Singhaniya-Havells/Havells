using Havells.D365.Entities.Survey.Request;
using Havells.D365.Entities.Survey.Response;

namespace Havells.D365.Services.Abstract
{
    public interface ISurveyRepository
    {
        SurveyDetail GetSurveyDetail(SurveyRequest request);
        UpdateSurveyResponse SubmitSurvey(SurveyUpdateRequest request);
    }
}
