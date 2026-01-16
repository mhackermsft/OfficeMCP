using OfficeMCP.Models;

namespace OfficeMCP.Services;

/// <summary>
/// Service interface for Word document operations.
/// Inherits all format-agnostic operations from IOfficeDocumentService.
/// Add Word-specific methods here if needed in the future.
/// </summary>
public interface IWordDocumentService : IOfficeDocumentService
{
    // All methods inherited from IOfficeDocumentService
    // Add Word-specific methods here if needed
}
