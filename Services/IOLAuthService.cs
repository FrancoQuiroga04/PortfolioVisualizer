using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace IOLPortfolio.Services
{
    public class IOLAuthService
    {
        private readonly HttpClient _http;
        private readonly string _username;
        private readonly string _password;
        public string AccessToken { get; private set; }
        public string RefreshToken { get; private set; }

        public IOLAuthService(HttpClient http, string username, string password)
        {
            _http = http;
            _username = username;
            _password = password;
        }

        public async Task<bool> LoginAsync()
        {
            var content = new StringContent(
                $"username={_username}&password={_password}&grant_type=password",
                Encoding.UTF8,
                "application/x-www-form-urlencoded"
            );

            var resp = await _http.PostAsync("token", content);
            if (!resp.IsSuccessStatusCode) return false;

            var json = await resp.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(json);

            AccessToken = doc.RootElement.GetProperty("access_token").GetString();
            RefreshToken = doc.RootElement.GetProperty("refresh_token").GetString();
            return true;
        }
    }
}
