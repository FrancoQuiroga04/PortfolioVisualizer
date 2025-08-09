using System.Net.Http;
using System.Text.Json;
using Microsoft.Extensions.Configuration;
using IOLPortfolio.Services;
using IOLPortfolio.Models;
using IOLPortfolio.Utils;

class Program
{
    static async Task Main()
    {
        var config = new ConfigurationBuilder()
            .SetBasePath(AppContext.BaseDirectory)
            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: false)
            .Build();

        var username = config["IOL:Username"] ?? Environment.GetEnvironmentVariable("IOL__Username");
        var password = config["IOL:Password"] ?? Environment.GetEnvironmentVariable("IOL__Password");

        if (string.IsNullOrWhiteSpace(username) || string.IsNullOrWhiteSpace(password))
        {
            Console.WriteLine("Faltan credenciales. Completá appsettings.json o setea IOL__Username / IOL__Password.");
            return;
        }

        using var http = new HttpClient { BaseAddress = new Uri("https://api.invertironline.com/") };

        try
        {
            var auth = new IOLAuthService(http, username, password);
            if (!await auth.LoginAsync())
            {
                Console.WriteLine("Error al iniciar sesión en IOL.");
                return;
            }

            Console.WriteLine("Login OK. Consultando portafolio...");

            var api = new IOLApiService(http, auth.AccessToken);
            var json = await api.GetPortfolioAsync();

            var rawPath = Path.Combine(AppContext.BaseDirectory, "portafolio_raw.json");
            await File.WriteAllTextAsync(rawPath, json);
            Console.WriteLine($"JSON guardado: {rawPath}");

            var lista = ParsePortfolio(json);
            if (lista.Count == 0)
            {
                Console.WriteLine("No se encontraron activos en el portafolio.");
                return;
            }

            var xlsxPath = Path.Combine(AppContext.BaseDirectory, "portafolio_iol.xlsx");
            ExcelExporter.Save(xlsxPath, lista);
            Console.WriteLine($"XLSX exportado: {xlsxPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Ocurrió un error:");
            Console.WriteLine(ex.Message);
        }
    }

    private static List<Portfolio> ParsePortfolio(string json)
    {
        var lista = new List<Portfolio>();

        using var doc = JsonDocument.Parse(json);
        if (!doc.RootElement.TryGetProperty("activos", out var activos) || activos.ValueKind != JsonValueKind.Array)
            return lista;

        foreach (var item in activos.EnumerateArray())
        {
            var titulo = item.GetProperty("titulo");

            lista.Add(new Portfolio
            {
                Simbolo = titulo.GetProperty("simbolo").GetString() ?? "",
                Descripcion = titulo.GetProperty("descripcion").GetString() ?? "",
                Tipo = titulo.GetProperty("tipo").GetString() ?? "",
                Mercado = titulo.GetProperty("mercado").GetString() ?? "",
                Moneda = titulo.GetProperty("moneda").GetString() ?? "",
                Cantidad = item.GetProperty("cantidad").GetDecimal(),
                UltimoPrecio = item.GetProperty("ultimoPrecio").GetDecimal(),
                Ppc = item.GetProperty("ppc").GetDecimal(),
                Valorizado = item.GetProperty("valorizado").GetDecimal(),
                GananciaDinero = item.GetProperty("gananciaDinero").GetDecimal(),
                GananciaPorcentaje = item.GetProperty("gananciaPorcentaje").GetDecimal(),
                VariacionDiaria = item.GetProperty("variacionDiaria").GetDecimal()
            });
        }

        return lista;
    }
}
