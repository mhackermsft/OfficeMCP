using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using ModelContextProtocol.Server;
using OfficeMCP.Services;
using OfficeMCP.Tools;

namespace OfficeMCP;

internal class Program
{
    static async Task Main(string[] args)
    {
        var builder = Host.CreateEmptyApplicationBuilder(settings: null);
        
        // Register format-specific services
        builder.Services.AddSingleton<IWordDocumentService, WordDocumentService>();
        builder.Services.AddSingleton<IExcelDocumentService, ExcelDocumentService>();
        builder.Services.AddSingleton<IPowerPointDocumentService, PowerPointDocumentService>();
        builder.Services.AddSingleton<IPdfDocumentService, PdfDocumentService>();
        
        // Register new services for Phase 0
        builder.Services.AddSingleton<IEncryptedDocumentHandler, EncryptedDocumentService>();
        
        // Configure MCP server with STDIO transport
        builder.Services
            .AddMcpServer(options =>
            {
                options.ServerInfo = new()
                {
                    Name = "OfficeMCP",
                    Version = "1.0.0"
                };
            })
            .WithStdioServerTransport()
            .WithToolsFromAssembly();

        await builder.Build().RunAsync();
    }
}

