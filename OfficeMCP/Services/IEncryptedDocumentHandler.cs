using OfficeMCP.Models;

namespace OfficeMCP.Services;

/// <summary>
/// Handles encrypted and sensitivity-labeled Office documents.
/// </summary>
public interface IEncryptedDocumentHandler
{
    /// <summary>
    /// Attempts to decrypt and read an encrypted document.
    /// </summary>
    Task<DocumentResult> DecryptAndReadAsync(string filePath, CancellationToken cancellationToken = default);

    /// <summary>
    /// Encrypts a document with specified protection.
    /// </summary>
    Task<DocumentResult> EncryptAndWriteAsync(string filePath, string content, CancellationToken cancellationToken = default);

    /// <summary>
    /// Gets label information from an encrypted document.
    /// </summary>
    Task<FileProtectionInfo> GetLabelInfoAsync(string filePath);

    /// <summary>
    /// Checks if the current user can access a protected document.
    /// </summary>
    Task<bool> CanUserAccessAsync(string filePath);
}

/// <summary>
/// Encryption configuration options.
/// </summary>
public record EncryptionConfig(
    bool EnableSensitivityLabelSupport = false,
    bool EnablePasswordProtectionPrompt = true,
    int DecryptedContentTimeoutMinutes = 60,
    string? AipUnifiedClientPath = null,
    bool UseSystemUserContext = true
);

/// <summary>
/// Extended file protection information including label details.
/// </summary>
public record FileProtectionInfo(
    bool IsProtected,
    bool IsEncrypted,
    bool MayHaveSensitivityLabel,
    bool IsValidOfficeFormat,
    string? ProtectionType,
    string? ErrorMessage,
    string? LabelId = null,
    string? LabelName = null,
    string[]? ContentMarkings = null,
    string? Encryption = null,
    string? CreatedBy = null,
    DateTime? CreatedDate = null,
    bool IsCompliant = false
);
