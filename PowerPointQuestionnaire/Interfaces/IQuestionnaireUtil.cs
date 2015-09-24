using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.RightsManagement;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointQuestionnaire.Model;

namespace PowerPointQuestionnaire.Interfaces
{
    public interface IQuestionnaireUtil
    {
        string Serialize(QuestionnaireModel questionnaire);
        QuestionnaireModel Deserialize(Slide slide);
        bool Check(Slide slide);
        void Mark(Slide slide, QuestionnaireModel questionnaire);
        void Unmark(Slide selectedSlide);
        Task<QuestionnaireModel> CreateAsync(QuestionnaireModel question);
        Task<QuestionnaireModel> UpdateAsync(QuestionnaireModel question);
        Task<QuestionnaireModel> DeleteAsync(string id);
        Task<QuestionnaireModel> GetAsync(string id);
        QuestionnaireModel Create(QuestionnaireModel question);
        QuestionnaireModel Update(QuestionnaireModel question);
        QuestionnaireModel Delete(string id);
        QuestionnaireModel Get(string id);
        bool Compare(QuestionnaireModel questionnaireGet, QuestionnaireModel questionnaire);
        Task<dynamic> UploadAsync(Slide slide);
        string GetSlideMd5(Slide selectedSlide);
    }
}
