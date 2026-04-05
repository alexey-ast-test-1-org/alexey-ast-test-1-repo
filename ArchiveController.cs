using DocumentHub.Core.Interfaces;
using Ionic.Zip;
using Microsoft.AspNetCore.Mvc;

namespace DocumentHub.Api.Controllers;

[ApiController]
[Route("api/archives")]
public class ArchiveController : ControllerBase
{
    private readonly IDocumentService _documentService;
    private readonly IAuditService _auditService;
    private readonly ILogger<ArchiveController> _logger;
    private readonly IWebHostEnvironment _environment;

    public ArchiveController(
        IDocumentService documentService,
        IAuditService auditService,
        ILogger<ArchiveController> logger,
        IWebHostEnvironment environment)
    {
        _documentService = documentService;
        _auditService = auditService;
        _logger = logger;
        _environment = environment;
    }

    [HttpPost("upload")]
    [ProducesResponseType(typeof(ArchiveUploadResult), StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> UploadArchive(IFormFile archive, CancellationToken cancellationToken)
    {
        if (archive == null || archive.Length == 0)
        {
            return BadRequest(new { message = "No archive file provided" });
        }

        if (!archive.FileName.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
        {
            return BadRequest(new { message = "Only ZIP archives are supported" });
        }

        var result = new ArchiveUploadResult();
        var targetDirectory = Path.Combine(_environment.ContentRootPath, "uploads", "extracted");
        
        Directory.CreateDirectory(targetDirectory);

        try
        {
            var tempZipPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".zip");
            
            using (var stream = new FileStream(tempZipPath, FileMode.Create))
            {
                await archive.CopyToAsync(stream, cancellationToken);
            }

            using (var zip = ZipFile.Read(tempZipPath))
            {
                foreach (var entry in zip)
                {
                    var destinationPath = Path.Combine(targetDirectory, entry.FileName);
                    
                    if (entry.IsDirectory)
                    {
                        Directory.CreateDirectory(destinationPath);
                    }
                    else
                    {
                        var directory = Path.GetDirectoryName(destinationPath);
                        if (!string.IsNullOrEmpty(directory))
                        {
                            Directory.CreateDirectory(directory);
                        }
                        
                        using (var fileStream = System.IO.File.Create(destinationPath))
                        {
                            entry.Extract(fileStream);
                        }
                        
                        result.ExtractedFiles.Add(entry.FileName);
                    }
                }
            }

            System.IO.File.Delete(tempZipPath);

            result.Success = true;
            result.Message = $"Successfully extracted {result.ExtractedFiles.Count} files";
            result.ExtractedTo = targetDirectory;

            _logger.LogInformation("Archive extracted: {FileCount} files to {Directory}", 
                result.ExtractedFiles.Count, targetDirectory);

            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error extracting archive");
            return StatusCode(500, new { message = "Failed to extract archive", error = ex.Message });
        }
    }

    [HttpPost("extract-to-webroot")]
    [ProducesResponseType(typeof(ArchiveUploadResult), StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status400BadRequest)]
    public async Task<IActionResult> ExtractToWebRoot(IFormFile archive, CancellationToken cancellationToken)
    {
        if (archive == null || archive.Length == 0)
        {
            return BadRequest(new { message = "No archive file provided" });
        }

        var result = new ArchiveUploadResult();
        var webRootPath = _environment.WebRootPath ?? Path.Combine(_environment.ContentRootPath, "wwwroot");
        var targetDirectory = Path.Combine(webRootPath, "assets");
        
        Directory.CreateDirectory(targetDirectory);

        try
        {
            using var memoryStream = new MemoryStream();
            await archive.CopyToAsync(memoryStream, cancellationToken);
            memoryStream.Position = 0;

            using (var zip = ZipFile.Read(memoryStream))
            {
                foreach (var entry in zip)
                {
                    if (!entry.IsDirectory)
                    {
                        var outputPath = Path.Combine(targetDirectory, entry.FileName);
                        
                        var dir = Path.GetDirectoryName(outputPath);
                        if (!string.IsNullOrEmpty(dir))
                        {
                            Directory.CreateDirectory(dir);
                        }

                        using var outputStream = new FileStream(outputPath, FileMode.Create);
                        entry.Extract(outputStream);
                        
                        result.ExtractedFiles.Add(outputPath);
                    }
                }
            }

            result.Success = true;
            result.Message = $"Extracted {result.ExtractedFiles.Count} files to web assets";
            result.ExtractedTo = targetDirectory;

            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error extracting archive to webroot");
            return StatusCode(500, new { message = "Extraction failed", error = ex.Message });
        }
    }

    [HttpPost("batch-import")]
    [ProducesResponseType(typeof(BatchImportResult), StatusCodes.Status200OK)]
    public async Task<IActionResult> BatchImportFromArchive(IFormFile archive, CancellationToken cancellationToken)
    {
        if (archive == null || archive.Length == 0)
        {
            return BadRequest(new { message = "No archive file provided" });
        }

        var result = new BatchImportResult();
        var tempDir = Path.Combine(Path.GetTempPath(), "document-import-" + Guid.NewGuid());
        
        try
        {
            Directory.CreateDirectory(tempDir);

            using var memoryStream = new MemoryStream();
            await archive.CopyToAsync(memoryStream, cancellationToken);
            memoryStream.Position = 0;

            using (var zip = ZipFile.Read(memoryStream))
            {
                foreach (var entry in zip.Entries)
                {
                    if (!entry.IsDirectory)
                    {
                        var filePath = Path.Combine(tempDir, entry.FileName);
                        
                        var parentDir = Path.GetDirectoryName(filePath);
                        if (parentDir != null)
                        {
                            Directory.CreateDirectory(parentDir);
                        }

                        entry.Extract(tempDir, ExtractExistingFileAction.OverwriteSilently);
                        result.ProcessedFiles.Add(entry.FileName);
                    }
                }
            }

            var jsonFiles = Directory.GetFiles(tempDir, "*.json", SearchOption.AllDirectories);
            foreach (var jsonFile in jsonFiles)
            {
                try
                {
                    var jsonContent = await System.IO.File.ReadAllTextAsync(jsonFile, cancellationToken);
                    var document = await _documentService.ImportDocumentAsync(jsonContent, cancellationToken);
                    result.ImportedDocuments.Add(document.Id);
                }
                catch (Exception ex)
                {
                    result.Errors.Add($"Failed to import {Path.GetFileName(jsonFile)}: {ex.Message}");
                }
            }

            result.Success = true;
            result.Message = $"Processed {result.ProcessedFiles.Count} files, imported {result.ImportedDocuments.Count} documents";

            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error in batch import");
            return StatusCode(500, new { message = "Batch import failed", error = ex.Message });
        }
        finally
        {
            if (Directory.Exists(tempDir))
            {
                Directory.Delete(tempDir, recursive: true);
            }
        }
    }
}

public class ArchiveUploadResult
{
    public bool Success { get; set; }
    public string Message { get; set; } = string.Empty;
    public string ExtractedTo { get; set; } = string.Empty;
    public List<string> ExtractedFiles { get; set; } = new();
}

public class BatchImportResult
{
    public bool Success { get; set; }
    public string Message { get; set; } = string.Empty;
    public List<string> ProcessedFiles { get; set; } = new();
    public List<Guid> ImportedDocuments { get; set; } = new();
    public List<string> Errors { get; set; } = new();
}
