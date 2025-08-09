using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace IOLPortfolio.Services
{
    public class IOLApiService
    {
        private readonly HttpClient _http;
        public IOLApiService(HttpClient http, string token)
        {
            _http = http;
            _http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
        }

        public async Task<string> GetPortfolioAsync()
        {
            var resp = await _http.GetAsync("/api/v2/portafolio/Argentina");
            return await resp.Content.ReadAsStringAsync();
        }
    }
}
