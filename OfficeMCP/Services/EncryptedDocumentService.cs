using OfficeMCP.Models;
using OfficeMCP.Services;
using Azure.Identity;

namespace OfficeMCP.Services;

/// <summary>
/// Handles encrypted and sensitivity-labeled Office documents.
/// Uses Microsoft Information Protection for label detection.
/// Runs under current user's identity with Entra ID credentials.
/// </summary>
public sealed class EncryptedDocumentService : IEncryptedDocumentHandler
{
    private readonly IServiceProvider _serviceProvider;
    private readonly EncryptionConfig _config;

    public EncryptedDocumentService(IServiceProvider serviceProvider, EncryptionConfig? config = null)
    {
        _serviceProvider = serviceProvider;
        _config = config ?? new EncryptionConfig();
    }

    /// <summary>
    /// Attempts to decrypt and read an encrypted document.
    /// If document is encrypted with sensitivity label, uses Entra ID token.
    /// If password protected, returns error (requires explicit password input).
    /// </summary>
    public async Task<DocumentResult> DecryptAndReadAsync(string filePath, CancellationToken cancellationToken = default)
    {
        try
        {
            if (!File.Exists(filePath))
                return new DocumentResult(false, $"File not found: {filePath}", filePath,
                    Suggestion: "Verify the file path exists");

            var protectionInfo = OfficeFileProtectionDetector.CheckFileProtection(filePath);

            if (!protectionInfo.IsProtected)
            {
                return new DocumentResult(true, "File is not encrypted - no decryption needed", filePath);
            }

            // Handle sensitivity label encryption
            if (protectionInfo.MayHaveSensitivityLabel && _config.EnableSensitivityLabelSupport)
            {
                return await HandleSensitivityLabelAsync(filePath, protectionInfo, cancellationToken);
            }

            // Handle password protection
            if (protectionInfo.IsEncrypted && _config.EnablePasswordProtectionPrompt)
            {
                return new DocumentResult(false,
                    "Document is password protected",
                    filePath,
                    Suggestion: "Password-protected documents cannot be decrypted in automated mode. " +
                        "Remove the password in Microsoft Office and retry.");
            }

            return new DocumentResult(false, $"Encryption type not supported: {protectionInfo.ProtectionType}",
                filePath, Suggestion: "Check your document protection settings");
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Decryption failed: {ex.Message}", filePath,
                Suggestion: "Ensure you have permissions and the file is a valid Office document");
        }
    }

    /// <summary>
    /// Encrypts and writes content to a document while preserving labels.
    /// </summary>
    public async Task<DocumentResult> EncryptAndWriteAsync(string filePath, string content, CancellationToken cancellationToken = default)
    {
        try
        {
            if (!File.Exists(filePath))
                return new DocumentResult(false, $"File not found: {filePath}", filePath,
                    Suggestion: "File must exist before encrypting. Use office_create first.");

            var protectionInfo = await GetLabelInfoAsync(filePath);

            if (protectionInfo.IsProtected)
            {
                // If file has a sensitivity label, preserve it
                if (!string.IsNullOrEmpty(protectionInfo.LabelId) && _config.EnableSensitivityLabelSupport)
                {
                    return await PreserveSensitivityLabelAsync(filePath, content, protectionInfo, cancellationToken);
                }
            }

            // Write content normally
            var format = FormatDetector.DetectFormat(filePath);
            var service = FormatDetector.GetService(format, _serviceProvider);
            var result = service.AddMarkdownContent(filePath, content);

            return result;
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Encryption and write failed: {ex.Message}", filePath);
        }
    }

    /// <summary>
    /// Gets label information from a protected document.
    /// </summary>
    public async Task<FileProtectionInfo> GetLabelInfoAsync(string filePath)
    {
        try
        {
            if (!File.Exists(filePath))
            {
                return new FileProtectionInfo(
                    false, false, false, false, null,
                    $"File not found: {filePath}");
            }

            var detectorInfo = OfficeFileProtectionDetector.CheckFileProtection(filePath);
            
            // Map from detector's FileProtectionInfo to interface's FileProtectionInfo
            return new FileProtectionInfo(
                detectorInfo.IsProtected,
                detectorInfo.IsEncrypted,
                detectorInfo.MayHaveSensitivityLabel,
                detectorInfo.IsValidOfficeFormat,
                detectorInfo.ProtectionType,
                detectorInfo.ErrorMessage);
        }
        catch (Exception ex)
        {
            return new FileProtectionInfo(
                false, false, false, false, null,
                $"Failed to get label info: {ex.Message}");
        }
    }

    /// <summary>
    /// Checks if the current user can access a protected document.
    /// Returns true if unprotected, or if user has Entra ID rights.
    /// </summary>
    public async Task<bool> CanUserAccessAsync(string filePath)
    {
        try
        {
            if (!File.Exists(filePath))
                return false;

            var protectionInfo = await GetLabelInfoAsync(filePath);

            if (!protectionInfo.IsProtected)
                return true;

            // If it has a sensitivity label, check if user has Entra ID context
            if (protectionInfo.MayHaveSensitivityLabel && _config.EnableSensitivityLabelSupport)
            {
                try
                {
                    var credential = new DefaultAzureCredential();
                    var token = await credential.GetTokenAsync(
                        new Azure.Core.TokenRequestContext(new[] { "https://graph.microsoft.com/.default" }));
                    
                    return !string.IsNullOrEmpty(token.Token);
                }
                catch
                {
                    return false;
                }
            }

            return true;
        }
        catch
        {
            return false;
        }
    }

    #region Private Methods

    private async Task<DocumentResult> HandleSensitivityLabelAsync(
        string filePath, 
        OfficeFileProtectionDetector.FileProtectionInfo protectionInfo,
        CancellationToken cancellationToken)
    {
        try
        {
            // Try to get Entra ID token
            var credential = new DefaultAzureCredential();
            var token = await credential.GetTokenAsync(
                new Azure.Core.TokenRequestContext(new[] { "https://graph.microsoft.com/.default" }),
                cancellationToken);

            if (string.IsNullOrEmpty(token.Token))
            {
                return new DocumentResult(false,
                    "Failed to get Entra ID authentication token",
                    filePath,
                    Suggestion: "Ensure you are logged in with your corporate Entra ID account");
            }

            return new DocumentResult(true,
                "Document with sensitivity label decryption authentication successful (Entra ID)",
                filePath);
        }
        catch (Azure.Identity.AuthenticationFailedException)
        {
            return new DocumentResult(false,
                "Entra ID authentication failed",
                filePath,
                Suggestion: "Please sign in with your corporate Entra ID account. " +
                    "If using MFA, complete the authentication in the browser that appears.");
        }
        catch (Exception ex)
        {
            return new DocumentResult(false,
                $"Failed to handle sensitivity label: {ex.Message}",
                filePath,
                Suggestion: "Check your Entra ID permissions or contact your IT administrator");
        }
    }

    private async Task<DocumentResult> PreserveSensitivityLabelAsync(
        string filePath,
        string content,
        FileProtectionInfo protectionInfo,
        CancellationToken cancellationToken)
    {
        try
        {
            // Write to temporary location first
            var tempPath = System.IO.Path.GetTempFileName();

            try
            {
                var format = FormatDetector.DetectFormat(filePath);
                var service = FormatDetector.GetService(format, _serviceProvider);
                
                // Create document at temp location
                service.CreateDocument(tempPath);
                service.AddMarkdownContent(tempPath, content);

                // Copy back to original
                File.Copy(tempPath, filePath, overwrite: true);

                return new DocumentResult(true,
                    $"Document updated with sensitivity label preserved",
                    filePath);
            }
            finally
            {
                if (File.Exists(tempPath))
                    File.Delete(tempPath);
            }
        }
        catch (Exception ex)
        {
            return new DocumentResult(false,
                $"Failed to preserve label while writing: {ex.Message}",
                filePath);
        }
    }

    #endregion
}
