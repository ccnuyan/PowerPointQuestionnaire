using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using PowerPointQuestionnaire.Interfaces;

namespace PowerPointQuestionnaire.Services
{
    public class AuthService : IAuthService
    {
        public static string Token;
        public static dynamic Me;

#if DEBUG
        private const string RestApiUrl = "http://localhost.:3000/api/";
#endif

#if RELEASE
        private const string RestApiUrl = "http://www.iccnu.net/api/";
#endif

        public async Task<bool> Authenticate(string username, string password)
        {
            try
            {
                var response = await GetTokenAsync(username, password);

                Token = (JsonConvert.DeserializeObject(response) as dynamic).access_token;

                var me = await GetMeAsync();

                Me = me;

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public HttpWebRequest AddToken(HttpWebRequest httpWebRequest)
        {
            httpWebRequest.Headers.Add("Authorization", "Bearer " + Token);
            return httpWebRequest;
        }

        public Task<string> GetTokenAsync(string username, string password)
        {
            return Task.Factory.StartNew(() => GetToken(username, password));
        }
        public Task<dynamic> GetMeAsync()
        {
            return Task.Factory.StartNew(() => GetMe());
        }

        public string GetToken(string username, string password)
        {
            var request = WebRequest.Create(RestApiUrl + "oauth2/token/") as HttpWebRequest;

            var body = string.Format("username={0}&password={1}&grant_type=password", username, password);

            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";

            request.Headers.Add("Authorization", "Basic cmVzb3VyY2Vfb3duZXJfdGVzdDpyZXNvdXJjZV9vd25lcl90ZXN0");

            var bts = Encoding.UTF8.GetBytes(body);
            request.ContentLength = bts.Length;
            using (var reqStream = request.GetRequestStream())
            {
                reqStream.Write(bts, 0, bts.Length);
                reqStream.Close();
            }

            using (var response = (HttpWebResponse)request.GetResponse())
            {
                using (var reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
                {
                    var responseData = reader.ReadToEnd();
                    return responseData;
                }
            }
        }
        public dynamic GetMe()
        {
            var request = AddToken(WebRequest.Create(RestApiUrl + "oauth2/me/") as HttpWebRequest);

            using (var response = (HttpWebResponse)request.GetResponse())
            {
                using (var reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
                {
                    var responseData = reader.ReadToEnd();
                    dynamic user = JsonConvert.DeserializeObject(responseData);
                    return user;
                }
            }
        }
    }
}
