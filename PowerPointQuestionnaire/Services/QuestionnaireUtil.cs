using PowerPointQuestionnaire.Interfaces;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using ClouDeveloper.Mime;
using Microsoft.Office.Interop.PowerPoint;
using Newtonsoft.Json;
using PowerPointQuestionnaire.Model;

namespace PowerPointQuestionnaire.Services
{
    public class QuestionnaireUtil : IQuestionnaireUtil
    {
        private readonly string _tempPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) +
                                  "PowerPointQuestionnaireImages/";

        public QuestionnaireUtil()
        {
            if (!Directory.Exists(_tempPath))
            {
                Directory.CreateDirectory(_tempPath);
            }
        }

#if DEBUG
        private const string RestApiUrl = "http://localhost.:3000/api/";
#else
        private const string RestApiUrl = "http://www.iccnu.net/api/";
#endif

        private const string RestQuestionnaireApiUrl = RestApiUrl+ "questionnaires/";
        private const string RestUploadApiUrl = RestApiUrl + "tempUpload/";

        private const string QuestionnaireMark = "wecourse.questionnaire";

        IAuthService _authService = new AuthService();

        public string Serialize(QuestionnaireModel questionnaire)
        {
            return JsonConvert.SerializeObject(questionnaire);
        }

        public QuestionnaireModel Deserialize(Slide slide)
        {
            var jsonstr = slide.Name.Substring(QuestionnaireMark.Length, slide.Name.Length - QuestionnaireMark.Length);

            var qm = JsonConvert.DeserializeObject<QuestionnaireModel>(jsonstr);

            return qm;
        }

        public bool Check(Slide slide)
        {
            return slide.Name.StartsWith(QuestionnaireMark);
        }

        public void Mark(Slide slide, QuestionnaireModel questionnaire)
        {
            slide.Name = QuestionnaireMark + Serialize(questionnaire);
        }

        public void Unmark(Slide slide)
        {
            slide.Name = Guid.NewGuid().ToString();
        }

        public Task<QuestionnaireModel> CreateAsync(QuestionnaireModel question)
        {
            return Task.Factory.StartNew(() => Create(question));
        }

        public QuestionnaireModel Create(QuestionnaireModel question)
        {
            if (!string.IsNullOrEmpty(question.id))
            {
                throw new Exception("try to create questionnaire with id");
            }

            var request = _authService.AddToken(WebRequest.Create(RestQuestionnaireApiUrl) as HttpWebRequest);

            request.Method = "POST";
            request.ContentType = "application/json";
            var bts = Encoding.UTF8.GetBytes(Serialize(question));
            request.ContentLength = bts.Length;
            using (var reqStream = request.GetRequestStream())
            {
                reqStream.Write(bts, 0, bts.Length);
                reqStream.Close();
            }

            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            {
                using (StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
                {
                    var responseData = reader.ReadToEnd();
                    var questionnareObject = JsonConvert.DeserializeObject<QuestionnaireModel>(responseData);
                    return questionnareObject;
                }
            }
        }

        public Task<QuestionnaireModel> UpdateAsync(QuestionnaireModel question)
        {
            return Task.Factory.StartNew(() => Update(question));
        }

        public QuestionnaireModel Update(QuestionnaireModel question)
        {
            if (string.IsNullOrEmpty(question.id))
            {
                throw new Exception("try to update questionnaire without id");
            }

            var request = _authService.AddToken(WebRequest.Create(RestQuestionnaireApiUrl + question.id) as HttpWebRequest);

            request.Method = "PUT";
            request.ContentType = "application/json";
            var bts = Encoding.UTF8.GetBytes(Serialize(question));
            request.ContentLength = bts.Length;
            using (var reqStream = request.GetRequestStream())
            {
                reqStream.Write(bts, 0, bts.Length);
                reqStream.Close();
            }

            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            {
                using (StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
                {
                    var responseData = reader.ReadToEnd();
                    var questionnareObject = JsonConvert.DeserializeObject<QuestionnaireModel>(responseData);
                    return questionnareObject;
                }
            }
        }

        public Task<QuestionnaireModel> DeleteAsync(string id)
        {
            return Task.Factory.StartNew(() => Delete(id));
        }

        public QuestionnaireModel Delete(string id)
        {
            var request = _authService.AddToken(WebRequest.Create(RestQuestionnaireApiUrl + id) as HttpWebRequest);

            request.Method = "DELETE";
            request.ContentType = "application/json";

            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            {
                using (StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
                {
                    var responseData = reader.ReadToEnd();
                    var questionnareObject = JsonConvert.DeserializeObject<QuestionnaireModel>(responseData);
                    return questionnareObject;
                }
            }
        }

        public Task<QuestionnaireModel> GetAsync(string id)
        {
            return Task.Factory.StartNew(() => Get(id));
        }

        public QuestionnaireModel Get(string id)
        {
            var request = _authService.AddToken(WebRequest.Create(RestQuestionnaireApiUrl + id) as HttpWebRequest);

            request.Method = "GET";
            request.ContentType = "application/json";

            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            {
                using (StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
                {
                    var responseData = reader.ReadToEnd();
                    var questionnareObject = JsonConvert.DeserializeObject<QuestionnaireModel>(responseData);
                    return questionnareObject;
                }
            }
        }

        public bool Compare(QuestionnaireModel q1, QuestionnaireModel q2)
        {
            if (q1.id != q2.id)
            {
                throw new Exception("comparing with different id");
            }

            return q1.choices == q2.choices;
        }

        public Task<dynamic> UploadAsync(Slide slide)
        {
            var tempFileString = Guid.NewGuid().ToString();

            var questionnaireImageFile = _tempPath + tempFileString + ".png";

            slide.Export(questionnaireImageFile, "png", 640, (int)(640 * slide.Master.Height / slide.Master.Width));

            return Task.Factory.StartNew(()=> UploadTempFile(questionnaireImageFile));
        }


        public dynamic UploadTempFile(string fileToBeUploaded)
        {
            //Console.WriteLine("the dir id:");
            //var rootdirid = Console.ReadLine();
            
            var fileInfo = new FileInfo(fileToBeUploaded);

            // Read file data
            FileStream fs = new FileStream(fileToBeUploaded, FileMode.Open, FileAccess.Read);
            byte[] data = new byte[fs.Length];
            fs.Read(data, 0, data.Length);
            fs.Close();

            // Generate post objects
            Dictionary<string, object> postParameters = new Dictionary<string, object>();

            var mime = MediaTypeNames.GetMediaTypeNames(fileInfo.Extension);
            postParameters.Add("file", new FormUploader.FileParameter(data, "test.txt", mime.First()));

            var request = _authService.AddToken(WebRequest.CreateHttp(RestUploadApiUrl));

            // Create request and receive response

            var userAgent = "starc wpf client";

            var webResponse = FormUploader.MultipartFormDataPost(request, userAgent, postParameters);

            // Process response
            var responseReader = new StreamReader(webResponse.GetResponseStream());

            var fullResponse = responseReader.ReadToEnd();

            webResponse.Close();

            return JsonConvert.DeserializeObject(fullResponse);
        }

        public string GetSlideMd5(Slide slide)
        {
            var questionnaire = Deserialize(slide);
            var questionnaireImageFile = _tempPath + questionnaire.id + ".png";
            if (File.Exists(questionnaireImageFile))
            {
                File.Delete(questionnaireImageFile);
            }
            slide.Export(questionnaireImageFile, "png", 640, (int)(640 * slide.Master.Height / slide.Master.Width));

            return GetFileMd5Code(questionnaireImageFile);
        }

        private string GetFileMd5Code(string filePath)
        {
            StringBuilder builder = new StringBuilder();
            using (var md5 = new MD5CryptoServiceProvider())
            {
                var tempPath = filePath + ".temp";
                File.Copy(filePath, tempPath, true);
                using (var fs = new FileStream(tempPath,FileMode.Open))
                {
                    byte[] bt = md5.ComputeHash(fs);
                    foreach (byte t in bt)
                    {
                        builder.Append(t.ToString("x2"));
                    }
                }
                File.Delete(tempPath);
            }
            return builder.ToString();
        }
    }
}
