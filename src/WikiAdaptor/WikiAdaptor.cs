using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace WikiAdaptor
{
    public class WikiAdaptor : IDisposable
    {
        Uri _baseAddress;
        HttpClient _client = new HttpClient();
        string _token = null;
        private bool disposedValue;

        public WikiAdaptor(string baseAddress)
        {
            _baseAddress = new Uri(baseAddress, UriKind.Absolute);
            _client.BaseAddress = _baseAddress;
        }
        public async void ClientLogin(string username, string password)
        {
            string token = await GetLoginToken();
            string escapedToken = Uri.EscapeDataString(token);
            string loginString = $"logintoken={escapedToken}&username={username}&password={password}";
            HttpContent loginContent = new StringContent(loginString, Encoding.UTF8, "application/x-www-form-urlencoded");
            var clientLoginResult = await _client.PostAsync($"api.php?action=clientlogin&format=json&loginreturnurl=https://petrowiki.org/", loginContent);
            if (!clientLoginResult.IsSuccessStatusCode)
                throw new HttpRequestException($"Login failed with {clientLoginResult.StatusCode}");
            _token = await GetCsrfToken();
        }
        public async void Login(string username, string password)
        {
            string token = await GetLoginToken();
            if(token != null)
            {
                string escapedToken = Uri.EscapeDataString(token);
                string loginString = $"lgtoken={escapedToken}&lgpassword={password}";
                HttpContent loginContent = new StringContent(loginString, Encoding.UTF8, "application/x-www-form-urlencoded");
                var loginResult = await _client.PostAsync($"api.php?action=login&format=json&lgname={username}", loginContent);
                if (!loginResult.IsSuccessStatusCode)
                    throw new HttpRequestException($"Login failed with {loginResult.StatusCode}");
                else
                {
                    var content = await loginResult.Content.ReadAsStringAsync();
                    JObject json = JsonConvert.DeserializeObject(content) as JObject;
                    var loginStatus = json["login"];
                    if(loginStatus != null)
                    {
                        string result = loginStatus.Value<string>("result");
                        if (string.Compare(result, "Success", true) != 0)
                        {
                            throw new ApplicationException($"Login failed. Result = {result}.");
                        }
                    }
                    else
                    {
                        throw new ApplicationException($"Login failed for unknown reason.");
                    }
                }
                _token = await GetCsrfToken();
            }
            else
            {
                throw new ApplicationException("Unable to get login token for wiki");
            }
        }
        public async Task<string> GetLoginToken()
        {
            var result = await _client.GetAsync("api.php?action=query&meta=tokens&format=json&type=login");
            if(result.IsSuccessStatusCode)
            {
                var content = await result.Content.ReadAsStringAsync();
                JObject json = JsonConvert.DeserializeObject(content) as JObject;
                return json["query"]["tokens"].Value<string>("logintoken");
            }
            return null;
        }
        public async Task<string> GetCsrfToken()
        {
            var result = await _client.GetAsync("api.php?action=query&meta=tokens&format=json");
            if (result.IsSuccessStatusCode)
            {
                var content = await result.Content.ReadAsStringAsync();
                JObject json = JsonConvert.DeserializeObject(content) as JObject;
                return json["query"]["tokens"].Value<string>("csrftoken");
            }
            return null;
        }
        private void CheckToken()
        {
            if(_token == null)
            {
                _token = GetCsrfToken().Result;
            }
        }
        private string GetEditstring(string title, string summary, bool overwrite = false)
        {
            var sb = new StringBuilder();
            sb.Append("api.php?action=edit");
            sb.Append("&format=json");
            sb.Append($"&title={Uri.EscapeDataString(title)}");
            sb.Append($"&summary={Uri.EscapeDataString(summary)}");
            sb.Append("&contentformat=text/x-wiki");
            if (!overwrite) sb.Append("&createonly");
            return sb.ToString();
        }
        public async Task CreatePage(string title, string summary, string content, bool overwrite = false)
        {
            CheckToken();
            var multipart = new MultipartFormDataContent(Guid.NewGuid().ToString());
            multipart.Add(new StringContent(_token), "token");
            multipart.Add(new StringContent(content), "text");
            var result = await _client.PostAsync(GetEditstring(title, summary, overwrite), multipart);
            if (!result.IsSuccessStatusCode)
            {
                throw new HttpRequestException(result.ReasonPhrase);
            }
        }
        public async Task CreatePage(string title, string summary, Stream content, bool overwrite = false)
        {
            CheckToken();
            var multipart = new MultipartFormDataContent(Guid.NewGuid().ToString());
            multipart.Add(new StringContent(_token), "token");
            multipart.Add(new StreamContent(content), "text");
            var result = await _client.PostAsync(GetEditstring(title, summary, overwrite), multipart);
            if (!result.IsSuccessStatusCode)
            {
                throw new HttpRequestException(result.ReasonPhrase);
            }
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    _client.Dispose();
                }

                disposedValue = true;
            }
        }


        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
