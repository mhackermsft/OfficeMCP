using System.IO.Compression;

namespace OfficeMCP.Services;

/// <summary>
/// Detects protection status of Office files including encryption and sensitivity labels.
/// </summary>
public static class OfficeFileProtectionDetector
{
    /// <summary>
    /// Information about file protection status.
    /// </summary>
    public record FileProtectionInfo(
        bool IsProtected,
        bool IsEncrypted,
        bool MayHaveSensitivityLabel,
        bool IsValidOfficeFormat,
        string ProtectionType,
        string? ErrorMessage = null
    );

    /// <summary>
    /// Checks if a file is protected/encrypted.
    /// </summary>
    public static FileProtectionInfo CheckFileProtection(string filePath)
    {
        if (!File.Exists(filePath))
        {
            return new FileProtectionInfo(
                false, false, false, false, "Unknown", $"File not found: {filePath}");
        }

        try
        {
            var extension = Path.GetExtension(filePath).ToLowerInvariant();
            
            // Check if it's an Office OpenXML file
            if (extension is ".docx" or ".xlsx" or ".pptx")
            {
                return CheckOpenXmlProtection(filePath);
            }
            
            // PDF files - check for encryption
            if (extension == ".pdf")
            {
                return CheckPdfProtection(filePath);
            }

            return new FileProtectionInfo(
                false, false, false, false, "Unknown", $"Unsupported file type: {extension}");
        }
        catch (Exception ex)
        {
            return new FileProtectionInfo(
                true, true, false, false, "Error", $"Could not check protection: {ex.Message}");
        }
    }

    private static FileProtectionInfo CheckOpenXmlProtection(string filePath)
    {
        try
        {
            // Try to open as ZIP to check if it's valid OpenXML
            using var archive = ZipFile.OpenRead(filePath);
            
            // Check for encryption markers in [Content_Types].xml
            var contentTypes = archive.GetEntry("[Content_Types].xml");
            if (contentTypes == null)
            {
                return new FileProtectionInfo(
                    true, true, false, false, "Encrypted", "File appears to be encrypted (invalid structure)");
            }

            // Check for Microsoft Information Protection markers
            var mipEntry = archive.Entries.FirstOrDefault(e => 
                e.FullName.Contains("LabelInfo", StringComparison.OrdinalIgnoreCase) ||
                e.FullName.Contains("protection", StringComparison.OrdinalIgnoreCase));
            
            bool mayHaveSensitivityLabel = mipEntry != null;

            return new FileProtectionInfo(
                mayHaveSensitivityLabel, false, mayHaveSensitivityLabel, true, 
                mayHaveSensitivityLabel ? "SensitivityLabel" : "None");
        }
        catch (InvalidDataException)
        {
            // File is not a valid ZIP/OpenXML - likely encrypted
            return new FileProtectionInfo(
                true, true, true, false, "Encrypted", 
                "File is encrypted and cannot be opened directly");
        }
        catch (Exception ex)
        {
            return new FileProtectionInfo(
                true, true, false, false, "Unknown", $"Error checking file: {ex.Message}");
        }
    }

    private static FileProtectionInfo CheckPdfProtection(string filePath)
    {
        try
        {
            using var stream = File.OpenRead(filePath);
            using var reader = new StreamReader(stream, leaveOpen: true);
            
            // Read first few KB to check for encryption markers
            var buffer = new char[4096];
            var read = reader.Read(buffer, 0, buffer.Length);
            var content = new string(buffer, 0, read);

            bool isEncrypted = content.Contains("/Encrypt", StringComparison.OrdinalIgnoreCase);
            
            return new FileProtectionInfo(
                isEncrypted, isEncrypted, false, true,
                isEncrypted ? "Password" : "None");
        }
        catch (Exception ex)
        {
            return new FileProtectionInfo(
                true, true, false, false, "Unknown", $"Error checking PDF: {ex.Message}");
        }
    }

    /// <summary>
    /// Gets a user-friendly error message for protected files.
    /// </summary>
    public static string GetProtectionErrorMessage(FileProtectionInfo info, string filePath)
    {
        if (!info.IsProtected)
            return string.Empty;

        var fileName = Path.GetFileName(filePath);

        return info.ProtectionType switch
        {
            "Encrypted" when info.MayHaveSensitivityLabel =>
                $"'{fileName}' is protected with a Microsoft sensitivity label. " +
                "You need appropriate permissions to access this file.",
            "Encrypted" =>
                $"'{fileName}' is encrypted or password-protected. " +
                "Remove the protection in the native application before using this tool.",
            "SensitivityLabel" =>
                $"'{fileName}' has a sensitivity label applied. " +
                "Content may be readable but some operations may be restricted.",
            "Password" =>
                $"'{fileName}' is password-protected. " +
                "Provide the password or remove protection in the native application.",
            _ => $"'{fileName}' has unknown protection: {info.ProtectionType}"
        };
    }
}
