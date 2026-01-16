using Microsoft.Extensions.DependencyInjection;
using OfficeMCP.Services;
using OfficeMCP.Tools;

namespace OfficeMCP.Tests;

/// <summary>
/// Shared test fixture that provides configured services for all tests.
/// </summary>
public class TestFixture : IDisposable
{
    public IServiceProvider ServiceProvider { get; }
    public string TestOutputDirectory { get; }
    private int _fileCounter = 0;

    public TestFixture()
    {
        // Create a unique test output directory for each test run
        TestOutputDirectory = Path.Combine(Path.GetTempPath(), "OfficeMCP_Tests", Guid.NewGuid().ToString("N")[..8]);
        Directory.CreateDirectory(TestOutputDirectory);

        // Configure services similar to Program.cs
        var services = new ServiceCollection();
        
        services.AddSingleton<IWordDocumentService, WordDocumentService>();
        services.AddSingleton<IExcelDocumentService, ExcelDocumentService>();
        services.AddSingleton<IPowerPointDocumentService, PowerPointDocumentService>();
        services.AddSingleton<IPdfDocumentService, PdfDocumentService>();
        services.AddSingleton<IEncryptedDocumentHandler, EncryptedDocumentService>();
        
        // Add the consolidated tools
        services.AddTransient<OfficeDocumentToolsConsolidated>();

        ServiceProvider = services.BuildServiceProvider();
    }

    /// <summary>
    /// Gets a unique test file path to avoid file conflicts between tests.
    /// </summary>
    public string GetTestFilePath(string fileName)
    {
        var counter = Interlocked.Increment(ref _fileCounter);
        var extension = Path.GetExtension(fileName);
        var nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
        return Path.Combine(TestOutputDirectory, $"{nameWithoutExt}_{counter}{extension}");
    }

    public void Dispose()
    {
        // Clean up test files
        try
        {
            if (Directory.Exists(TestOutputDirectory))
            {
                Directory.Delete(TestOutputDirectory, recursive: true);
            }
        }
        catch
        {
            // Ignore cleanup errors
        }
        GC.SuppressFinalize(this);
    }
}

/// <summary>
/// Collection definition for sharing the test fixture across test classes.
/// </summary>
[CollectionDefinition("Office Tests")]
public class OfficeTestCollection : ICollectionFixture<TestFixture>
{
}
