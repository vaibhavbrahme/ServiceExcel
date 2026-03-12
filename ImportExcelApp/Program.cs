using System;
using System.IO;
using System.Net;
using System.Text.Json;
using System.Threading.Tasks;
using ImportExcelApp.Models;
using ImportExcelApp.Services;

// simple self-hosted HTTP listener so the app can run on .NET Framework 4.8
public class Program
{
    private const string Prefix = "http://localhost:5055/";

    public static async Task Main(string[] args)
    {
        using var listener = new HttpListener();
        listener.Prefixes.Add(Prefix);
            try
            {
                listener.Start();
                Console.WriteLine($"Listening on {Prefix}");
            }
            catch (HttpListenerException hlex)
            {
                Console.WriteLine($"Failed to start listener on {Prefix}: {hlex.Message}");
                Console.WriteLine("This usually means another process has already reserved the URL. Try stopping that process or use a different port in config.");
                return;
            }

            // service instance used for all requests
            var excelService = new ExcelService();

        // configuration file logic removed; sheet name is taken directly from the request body

        while (true)
        {
            HttpListenerContext context;
            try
            {
                context = await listener.GetContextAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Listener error: " + ex);
                // short pause then continue listening
                await Task.Delay(1000);
                continue;
            }

            // await the handler so any exceptions propagate here for logging
            try
            {
                await HandleRequestAsync(context, excelService);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unhandled error processing request: " + ex);
            }
        }
    }

    private static async Task HandleRequestAsync(HttpListenerContext context, IExcelService excelService)
    {
        try
        {
            if (context.Request.HttpMethod == "POST" &&
                context.Request.Url.AbsolutePath.Equals("/api/excel/import", StringComparison.OrdinalIgnoreCase))
            {
                using var reader = new StreamReader(context.Request.InputStream);
                var body = await reader.ReadToEndAsync();
                Console.WriteLine("Received body: " + body);

                var options = new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                };
                var request = JsonSerializer.Deserialize<ExcelImportRequest>(body, options);

                if (request == null)
                {
                    Console.WriteLine("Warning: request deserialized to null");
                    context.Response.StatusCode = 400;
                    using (var writer = new StreamWriter(context.Response.OutputStream))
                    {
                        writer.Write("{\"success\":false,\"message\":\"Invalid request payload\"}");
                    }
                    return;
                }

    
                var result = await excelService.ImportExcelAsync(request);

                context.Response.StatusCode = result.Success ? 200 : 400;
                context.Response.ContentType = "application/json";
                await JsonSerializer.SerializeAsync(context.Response.OutputStream, result);
            }
            else
            {
                context.Response.StatusCode = 404;
            }
        }
        catch (Exception ex)
        {
            context.Response.StatusCode = 500;
            using var writer = new StreamWriter(context.Response.OutputStream);
            await writer.WriteAsync($"{{\"success\":false,\"message\":\"{ex.Message}\"}}");
        }
        finally
        {
            context.Response.OutputStream.Close();
        }
    }
}
