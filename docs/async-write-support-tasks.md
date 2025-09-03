# Async Write Support Implementation Tasks

## Overview
This document contains the complete implementation plan for adding asynchronous write support to NPOI's XSSFWorkbook class. The implementation is divided into three phases, each building upon the previous one.

---

## Phase 1: Core Infrastructure

### Overview
This phase establishes the foundational async infrastructure for NPOI's write operations. The focus is on creating the core async methods without modifying existing synchronous functionality.

### Tasks

#### 1. Interface Extension (IWorkbook)
**File:** `/main/SS/UserModel/Workbook.cs`

**Action:** Add async method signature to IWorkbook interface
```csharp
/// <summary>
/// Write the workbook to the specified stream asynchronously, and optionally leave the stream open without closing it.
/// </summary>
/// <param name="stream">the stream you wish to write the workbook to</param>
/// <param name="leaveOpen">leave stream open or not</param>
/// <param name="cancellationToken">cancellation token to observe during the async operation</param>
Task WriteAsync(Stream stream, bool leaveOpen = false, CancellationToken cancellationToken = default);
```

#### 2. XSSFWorkbook.WriteAsync Implementation
**File:** `/ooxml/XSSF/UserModel/XSSFWorkbook.cs`

**Action:** Implement the main async entry point
```csharp
public async Task WriteAsync(Stream stream, bool leaveOpen = false, CancellationToken cancellationToken = default)
{
    bool? originalValue = null;
    if (Package is ZipPackage package)
    {
        originalValue = package.IsExternalStream;
        package.IsExternalStream = leaveOpen;
    }
    
    await base.WriteAsync(stream, cancellationToken).ConfigureAwait(false);
    
    if (originalValue.HasValue && Package is ZipPackage zipPackage)
    {
        zipPackage.IsExternalStream = originalValue.Value;
    }
}
```

#### 3. POIXMLDocument.WriteAsync Implementation  
**File:** `/ooxml/POIXMLDocument.cs`

**Action:** Create async version of Write method
```csharp
public async Task WriteAsync(Stream stream, CancellationToken cancellationToken = default)
{
    OPCPackage pkg = Package;
    if (pkg == null)
    {
        throw new IOException("Cannot write data, document seems to have been closed already");
    }
    
    if (!this.GetProperties().CustomProperties.Contains("Generator"))
        this.GetProperties().CustomProperties.AddProperty("Generator", "NPOI");
    if (!this.GetProperties().CustomProperties.Contains("Generator Version"))
        this.GetProperties().CustomProperties.AddProperty("Generator Version", Assembly.GetExecutingAssembly().GetName().Version.ToString(3));
        
    //force all children to commit their changes into the underlying OOXML Package
    List<PackagePart> context = new List<PackagePart>();
    await OnSaveAsync(context, cancellationToken).ConfigureAwait(false);
    context.Clear();

    //save extended and custom properties
    await GetProperties().CommitAsync(cancellationToken).ConfigureAwait(false);

    await pkg.SaveAsync(stream, cancellationToken).ConfigureAwait(false);
}

protected virtual async Task OnSaveAsync(List<PackagePart> context, CancellationToken cancellationToken)
{
    // Default implementation calls synchronous OnSave for backward compatibility
    // Derived classes can override for true async behavior
    await Task.Run(() => OnSave(context), cancellationToken).ConfigureAwait(false);
}
```

#### 4. OPCPackage.SaveAsync Infrastructure
**File:** `/openxml4Net/OPC/OPCPackage.cs`

**Action:** Add abstract SaveAsync method
```csharp
public async Task SaveAsync(Stream outputStream, CancellationToken cancellationToken = default)
{
    ThrowExceptionIfReadOnly();
    await SaveImplAsync(outputStream, cancellationToken).ConfigureAwait(false);
}

protected abstract Task SaveImplAsync(Stream outputStream, CancellationToken cancellationToken = default);
```

#### 5. ZipPackage.SaveImplAsync Implementation
**File:** `/openxml4Net/OPC/ZipPackage.cs`

**Action:** Implement async ZIP package saving
```csharp
protected override async Task SaveImplAsync(Stream outputStream, CancellationToken cancellationToken = default)
{
    // Check that the document was open in write mode
    ThrowExceptionIfReadOnly();
    
    using var zos = outputStream is ZipOutputStream zipStream ? zipStream : new ZipOutputStream(outputStream);
    zos.UseZip64 = UseZip64.Off;
    
    // Add core properties if missing
    if (this.GetPartsByRelationshipType(PackageRelationshipTypes.CORE_PROPERTIES).Count == 0 &&
        this.GetPartsByRelationshipType(PackageRelationshipTypes.CORE_PROPERTIES_ECMA376).Count == 0)
    {
        logger.Log(POILogger.DEBUG, "Save core properties part");
        GetPackageProperties();
        AddPackagePart(this.packageProperties);
        this.relationships.AddRelationship(this.packageProperties
                .PartName.URI, TargetMode.Internal,
                PackageRelationshipTypes.CORE_PROPERTIES, null);
    }
    
    // Save package relationships
    await WritePackageRelationshipsAsync(zos, cancellationToken).ConfigureAwait(false);
    
    // Save content types
    await WriteContentTypesAsync(zos, cancellationToken).ConfigureAwait(false);
    
    // Save package parts
    await WritePartsAsync(zos, cancellationToken).ConfigureAwait(false);
    
    await zos.FlushAsync(cancellationToken).ConfigureAwait(false);
}
```

### Phase 1 Testing Requirements

#### Unit Tests
- **Test File:** `/testcases/ooxml/XSSF/UserModel/TestXSSFWorkbookAsync.cs`
- Test basic WriteAsync functionality
- Verify leaveOpen parameter behavior
- Test cancellation token handling
- Ensure identical output between sync/async versions

#### Integration Points
- Verify no breaking changes to existing sync methods
- Test with real workbook data
- Memory usage validation

### Phase 1 Success Criteria
1. All async methods compile without errors
2. Existing synchronous tests continue to pass
3. Basic async write operations work correctly
4. Proper exception handling and cancellation support
5. No memory leaks or resource disposal issues

---

## Phase 2: XML Serialization

### Overview
This phase focuses on optimizing XML serialization operations for true async performance. Building on Phase 1's infrastructure, we'll replace async-over-sync patterns with native async XML writing.

### Prerequisites
- Phase 1 core infrastructure completed
- Basic async write pipeline functional

### Tasks

#### 1. Async XML Helper Infrastructure
**File:** `/OpenXmlFormats/XmlHelper.cs`

**Action:** Add async XML writing utilities
```csharp
public static async Task WriteAttributeAsync(TextWriter sw, string attributeName, string value, CancellationToken cancellationToken = default)
{
    if (value != null)
    {
        await sw.WriteAsync($" {attributeName}=\"").ConfigureAwait(false);
        await sw.WriteAsync(XmlConvert.EncodeName(value)).ConfigureAwait(false);
        await sw.WriteAsync("\"").ConfigureAwait(false);
    }
}

public static async Task WriteElementAsync(TextWriter sw, string elementName, string value, CancellationToken cancellationToken = default)
{
    if (value != null)
    {
        await sw.WriteAsync($"<{elementName}>").ConfigureAwait(false);
        await sw.WriteAsync(XmlConvert.EncodeName(value)).ConfigureAwait(false);
        await sw.WriteAsync($"</{elementName}>").ConfigureAwait(false);
    }
}
```

#### 2. Workbook XML Async Serialization
**File:** `/OpenXmlFormats/Spreadsheet/Workbook.cs`

**Action:** Add async Write methods to workbook components
```csharp
internal async Task WriteAsync(TextWriter sw, string nodeName, CancellationToken cancellationToken = default)
{
    await sw.WriteAsync(string.Format("<{0}", nodeName)).ConfigureAwait(false);
    await XmlHelper.WriteAttributeAsync(sw, "appName", this.appName, cancellationToken).ConfigureAwait(false);
    await XmlHelper.WriteAttributeAsync(sw, "lastEdited", this.lastEdited, cancellationToken).ConfigureAwait(false);
    // ... continue for all attributes
    await sw.WriteAsync(">").ConfigureAwait(false);
    await sw.WriteAsync(string.Format("</{0}>", nodeName)).ConfigureAwait(false);
}
```

#### 3. Shared Strings Table Async Writing  
**File:** `/ooxml/XSSF/Model/SharedStringsTable.cs`

**Action:** Implement async serialization for shared strings
```csharp
public async Task WriteToAsync(Stream outputStream, CancellationToken cancellationToken = default)
{
    using var writer = new StreamWriter(outputStream, Encoding.UTF8, leaveOpen: true);
    await writer.WriteAsync("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>").ConfigureAwait(false);
    await writer.WriteAsync("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"").ConfigureAwait(false);
    await writer.WriteAsync($" count=\"{strings.Count}\" uniqueCount=\"{strings.Count}\">").ConfigureAwait(false);
    
    foreach (var str in strings)
    {
        await WriteSharedStringAsync(writer, str, cancellationToken).ConfigureAwait(false);
    }
    
    await writer.WriteAsync("</sst>").ConfigureAwait(false);
    await writer.FlushAsync().ConfigureAwait(false);
}

private async Task WriteSharedStringAsync(TextWriter writer, CT_Rst rst, CancellationToken cancellationToken)
{
    await writer.WriteAsync("<si>").ConfigureAwait(false);
    await rst.WriteAsync(writer, "si", cancellationToken).ConfigureAwait(false);
    await writer.WriteAsync("</si>").ConfigureAwait(false);
}
```

#### 4. Styles Table Async Writing
**File:** `/ooxml/XSSF/Model/StylesTable.cs`

**Action:** Implement async serialization for styles
```csharp
public async Task WriteToAsync(Stream outputStream, CancellationToken cancellationToken = default)
{
    using var writer = new StreamWriter(outputStream, Encoding.UTF8, leaveOpen: true);
    await writer.WriteAsync("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>").ConfigureAwait(false);
    await doc.WriteAsync(writer, cancellationToken).ConfigureAwait(false);
    await writer.FlushAsync().ConfigureAwait(false);
}
```

#### 5. Worksheet Async Writing
**File:** `/ooxml/XSSF/UserModel/XSSFSheet.cs`

**Action:** Add async serialization for worksheet data
```csharp
public async Task WriteToAsync(Stream outputStream, CancellationToken cancellationToken = default)
{
    using var writer = new StreamWriter(outputStream, Encoding.UTF8, leaveOpen: true);
    await writer.WriteAsync("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>").ConfigureAwait(false);
    await worksheet.WriteAsync(writer, cancellationToken).ConfigureAwait(false);
    await writer.FlushAsync().ConfigureAwait(false);
}
```

#### 6. Enhanced ZipPackage Async Operations
**File:** `/openxml4Net/OPC/ZipPackage.cs`

**Action:** Optimize async ZIP entry writing with true async XML serialization
```csharp
private async Task WritePackageRelationshipsAsync(ZipOutputStream zos, CancellationToken cancellationToken)
{
    if (relationships.Size > 0)
    {
        var entry = new ZipEntry(PackagingURIHelper.PACKAGE_RELATIONSHIPS_ROOT_URI.ToString().Substring(1));
        zos.PutNextEntry(entry);
        
        using var writer = new StreamWriter(zos, Encoding.UTF8, leaveOpen: true);
        await relationships.WriteAsync(writer, cancellationToken).ConfigureAwait(false);
        await writer.FlushAsync().ConfigureAwait(false);
        
        zos.CloseEntry();
    }
}

private async Task WriteContentTypesAsync(ZipOutputStream zos, CancellationToken cancellationToken)
{
    var entry = new ZipEntry("[Content_Types].xml");
    zos.PutNextEntry(entry);
    
    using var writer = new StreamWriter(zos, Encoding.UTF8, leaveOpen: true);
    await contentTypeManager.WriteAsync(writer, cancellationToken).ConfigureAwait(false);
    await writer.FlushAsync().ConfigureAwait(false);
    
    zos.CloseEntry();
}

private async Task WritePartsAsync(ZipOutputStream zos, CancellationToken cancellationToken)
{
    foreach (PackagePart part in getParts())
    {
        cancellationToken.ThrowIfCancellationRequested();
        
        var entry = new ZipEntry(part.PartName.ToString().Substring(1));
        zos.PutNextEntry(entry);
        
        if (part.GetContentType().Contains("xml"))
        {
            await WriteXmlPartAsync(part, zos, cancellationToken).ConfigureAwait(false);
        }
        else
        {
            await WriteBinaryPartAsync(part, zos, cancellationToken).ConfigureAwait(false);
        }
        
        zos.CloseEntry();
    }
}
```

#### 7. XSSFWorkbook OnSaveAsync Override
**File:** `/ooxml/XSSF/UserModel/XSSFWorkbook.cs`

**Action:** Override OnSaveAsync for optimized worksheet saving
```csharp
protected override async Task OnSaveAsync(List<PackagePart> context, CancellationToken cancellationToken)
{
    // Save shared strings table
    if (sharedStringSource != null && sharedStringSource.Count > 0)
    {
        var part = GetPackagePart(PackageRelationshipTypes.SHARED_STRINGS);
        await sharedStringSource.WriteToAsync(part.GetOutputStream(), cancellationToken).ConfigureAwait(false);
    }
    
    // Save styles
    if (stylesSource != null)
    {
        var part = GetPackagePart(PackageRelationshipTypes.STYLES);
        await stylesSource.WriteToAsync(part.GetOutputStream(), cancellationToken).ConfigureAwait(false);
    }
    
    // Save worksheets
    foreach (var sheet in sheets)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var part = sheet.GetPackagePart();
        await sheet.WriteToAsync(part.GetOutputStream(), cancellationToken).ConfigureAwait(false);
    }
    
    // Save workbook XML
    await WriteWorkbookXmlAsync(cancellationToken).ConfigureAwait(false);
}
```

### Phase 2 Performance Optimizations

#### Streaming XML Writing
- Use `StreamWriter` with async methods instead of building strings in memory
- Write XML content incrementally for large worksheets
- Implement buffering strategies for optimal I/O performance

#### Memory Management
- Dispose writers properly in async contexts
- Use `leaveOpen: true` for intermediate streams
- Implement async enumerable patterns for large collections

### Phase 2 Testing Requirements

#### Unit Tests
- **Test File:** `/testcases/ooxml/XSSF/UserModel/TestXSSFWorkbookAsyncPhase2.cs`
- Test async XML serialization correctness
- Verify performance improvements over async-over-sync
- Test cancellation during XML writing
- Memory usage tests for large workbooks

#### Performance Tests
- Benchmark async vs sync XML serialization
- Memory allocation profiling
- Throughput tests with large datasets

### Phase 2 Success Criteria
1. All XML serialization operations use native async I/O
2. Measurable performance improvement over Phase 1
3. Memory usage remains stable or improves
4. Proper cancellation handling during XML operations
5. Output remains bit-identical to synchronous version

---

## Phase 3: Binary Content & Polish

### Overview
This final phase optimizes binary content handling and adds polish to the async write implementation. Focus areas include images, fonts, embedded objects, and overall system optimization.

### Prerequisites
- Phase 1 core infrastructure completed
- Phase 2 XML serialization optimized
- All basic async functionality working correctly

### Tasks

#### 1. Async Binary Data Handling
**File:** `/openxml4Net/OPC/ZipPackage.cs`

**Action:** Optimize binary content copying with native async operations
```csharp
private async Task WriteBinaryPartAsync(PackagePart part, Stream outputStream, CancellationToken cancellationToken)
{
    using var partStream = part.GetInputStream();
    await partStream.CopyToAsync(outputStream, 81920, cancellationToken).ConfigureAwait(false); // 80KB buffer
}

private async Task WriteXmlPartAsync(PackagePart part, Stream outputStream, CancellationToken cancellationToken)
{
    // For XML parts that haven't been optimized yet, use smart detection
    if (part is IAsyncWritable asyncWritable)
    {
        await asyncWritable.WriteToAsync(outputStream, cancellationToken).ConfigureAwait(false);
    }
    else
    {
        // Fallback to stream copying for non-optimized XML parts
        using var partStream = part.GetInputStream();
        await partStream.CopyToAsync(outputStream, 81920, cancellationToken).ConfigureAwait(false);
    }
}
```

#### 2. Image Handling Optimization
**File:** `/ooxml/XSSF/UserModel/XSSFPicture.cs`

**Action:** Implement async image data processing
```csharp
public async Task WriteImageDataAsync(Stream outputStream, CancellationToken cancellationToken = default)
{
    var pictureData = GetPictureData();
    if (pictureData?.Data != null)
    {
        await outputStream.WriteAsync(pictureData.Data, 0, pictureData.Data.Length, cancellationToken).ConfigureAwait(false);
    }
}
```

#### 3. Font Resource Async Loading
**File:** `/main/SS/UserModel/FontFamily.cs` (if exists)

**Action:** Optimize font resource handling for async operations
```csharp
public static async Task<byte[]> LoadFontDataAsync(string fontPath, CancellationToken cancellationToken = default)
{
    if (File.Exists(fontPath))
    {
        return await File.ReadAllBytesAsync(fontPath, cancellationToken).ConfigureAwait(false);
    }
    return null;
}
```

#### 4. Progress Reporting Infrastructure
**File:** `/ooxml/XSSF/UserModel/XSSFWorkbook.cs`

**Action:** Add optional progress reporting for large workbooks
```csharp
public async Task WriteAsync(Stream stream, bool leaveOpen = false, IProgress<WriteProgress> progress = null, CancellationToken cancellationToken = default)
{
    var progressReporter = new WriteProgressReporter(progress);
    
    bool? originalValue = null;
    if (Package is ZipPackage package)
    {
        originalValue = package.IsExternalStream;
        package.IsExternalStream = leaveOpen;
    }
    
    try
    {
        progressReporter.Report("Starting write operation", 0);
        await base.WriteAsync(stream, progressReporter, cancellationToken).ConfigureAwait(false);
        progressReporter.Report("Write completed", 100);
    }
    finally
    {
        if (originalValue.HasValue && Package is ZipPackage zipPackage)
        {
            zipPackage.IsExternalStream = originalValue.Value;
        }
    }
}

public class WriteProgress
{
    public string CurrentOperation { get; set; }
    public int PercentComplete { get; set; }
    public long BytesWritten { get; set; }
    public TimeSpan ElapsedTime { get; set; }
}
```

#### 5. Memory Pool Integration
**File:** `/ooxml/XSSF/Streaming/SXSSFWorkbook.cs`

**Action:** Optimize memory usage with ArrayPool for large operations
```csharp
private static readonly ArrayPool<byte> BufferPool = ArrayPool<byte>.Shared;

public async Task WriteAsync(Stream stream, bool leaveOpen = false, CancellationToken cancellationToken = default)
{
    byte[] buffer = BufferPool.Rent(81920); // 80KB buffer
    try
    {
        // Use pooled buffer for operations
        await base.WriteAsync(stream, leaveOpen, cancellationToken).ConfigureAwait(false);
    }
    finally
    {
        BufferPool.Return(buffer);
    }
}
```

#### 6. Enhanced Error Handling & Diagnostics
**File:** `/ooxml/POIXMLDocument.cs`

**Action:** Add comprehensive error handling and diagnostics
```csharp
public async Task WriteAsync(Stream stream, CancellationToken cancellationToken = default)
{
    var stopwatch = Stopwatch.StartNew();
    
    try
    {
        OPCPackage pkg = Package;
        if (pkg == null)
        {
            throw new IOException("Cannot write data, document seems to have been closed already");
        }
        
        logger.Log(POILogger.DEBUG, "Starting async write operation");
        
        // Set document properties
        EnsureDocumentProperties();
        
        // Save all parts
        List<PackagePart> context = new List<PackagePart>();
        await OnSaveAsync(context, cancellationToken).ConfigureAwait(false);
        context.Clear();

        // Commit properties and save package
        await GetProperties().CommitAsync(cancellationToken).ConfigureAwait(false);
        await pkg.SaveAsync(stream, cancellationToken).ConfigureAwait(false);
        
        logger.Log(POILogger.DEBUG, $"Async write completed in {stopwatch.ElapsedMilliseconds}ms");
    }
    catch (OperationCanceledException)
    {
        logger.Log(POILogger.INFO, $"Write operation cancelled after {stopwatch.ElapsedMilliseconds}ms");
        throw;
    }
    catch (Exception ex)
    {
        logger.Log(POILogger.ERROR, $"Write operation failed after {stopwatch.ElapsedMilliseconds}ms: {ex.Message}");
        throw;
    }
}
```

#### 7. Concurrent Processing Support
**File:** `/ooxml/XSSF/UserModel/XSSFWorkbook.cs`

**Action:** Enable concurrent processing of independent worksheet parts
```csharp
protected override async Task OnSaveAsync(List<PackagePart> context, CancellationToken cancellationToken)
{
    var tasks = new List<Task>();
    
    // Process independent components concurrently
    if (sharedStringSource != null && sharedStringSource.Count > 0)
    {
        tasks.Add(SaveSharedStringsAsync(cancellationToken));
    }
    
    if (stylesSource != null)
    {
        tasks.Add(SaveStylesAsync(cancellationToken));
    }
    
    // Wait for shared resources to complete first
    if (tasks.Count > 0)
    {
        await Task.WhenAll(tasks).ConfigureAwait(false);
        tasks.Clear();
    }
    
    // Then process worksheets (which may depend on shared resources)
    var worksheetTasks = sheets.Select(sheet => SaveWorksheetAsync(sheet, cancellationToken));
    await Task.WhenAll(worksheetTasks).ConfigureAwait(false);
    
    // Finally save workbook XML
    await SaveWorkbookXmlAsync(cancellationToken).ConfigureAwait(false);
}
```

#### 8. Streaming API Integration
**File:** `/ooxml/XSSF/Streaming/SXSSFWorkbook.cs`

**Action:** Optimize SXSSF streaming workbook for async operations
```csharp
public override async Task WriteAsync(Stream stream, bool leaveOpen = false, CancellationToken cancellationToken = default)
{
    await FlushAllSheetsAsync(cancellationToken).ConfigureAwait(false);
    await base.WriteAsync(stream, leaveOpen, cancellationToken).ConfigureAwait(false);
}

private async Task FlushAllSheetsAsync(CancellationToken cancellationToken)
{
    var flushTasks = _sxssfSheets.Values.Select(sheet => sheet.FlushRowsAsync(cancellationToken));
    await Task.WhenAll(flushTasks).ConfigureAwait(false);
}
```

### Phase 3 Performance Optimizations

#### I/O Optimization
- Use larger buffer sizes for binary data (80KB+)
- Implement read-ahead buffering for large embedded content
- Optimize ZIP compression settings for async operations

#### Memory Management  
- Integrate `System.Buffers.ArrayPool` for temporary allocations
- Implement object pooling for frequently created objects
- Use `Memory<T>` and `ReadOnlyMemory<T>` patterns where applicable

#### Concurrency
- Process independent worksheet parts concurrently
- Pipeline ZIP entry writing with content generation
- Use semaphores to limit concurrent I/O operations

### Phase 3 Testing Requirements

#### Integration Tests
- **Test File:** `/testcases/ooxml/XSSF/UserModel/TestXSSFWorkbookAsyncComplete.cs`
- End-to-end async write testing with complex workbooks
- Performance regression testing
- Memory usage validation under load
- Cancellation during various phases

#### Performance Tests  
- Large workbook benchmarks (>100MB)
- Memory pressure testing
- Concurrent write operations
- Progress reporting accuracy

#### Real-world Scenarios
- Workbooks with many embedded images
- Complex formatting and styles
- Large datasets with streaming
- Network stream scenarios

### Phase 3 Success Criteria
1. Binary content handling optimized for async operations
2. Memory usage improved or maintained under all scenarios  
3. Cancellation works reliably during all phases
4. Progress reporting provides accurate feedback
5. Performance matches or exceeds synchronous operations
6. No resource leaks or disposal issues
7. Concurrent processing improves overall throughput

---

## Testing Strategy

### Test Categories

#### 1. Unit Tests

##### Basic Functionality Tests
**Location:** `/testcases/ooxml/XSSF/UserModel/TestXSSFWorkbookAsync.cs`

```csharp
[Test]
public async Task TestBasicWriteAsync()
{
    using var workbook = new XSSFWorkbook();
    var sheet = workbook.CreateSheet("Test");
    var row = sheet.CreateRow(0);
    var cell = row.CreateCell(0);
    cell.SetCellValue("Hello World");
    
    using var asyncStream = new MemoryStream();
    using var syncStream = new MemoryStream();
    
    // Write both async and sync versions
    await workbook.WriteAsync(asyncStream);
    workbook.Write(syncStream);
    
    // Verify identical output
    Assert.AreEqual(syncStream.ToArray(), asyncStream.ToArray());
}

[Test]
public async Task TestWriteAsyncWithLeaveOpen()
{
    using var workbook = new XSSFWorkbook();
    using var stream = new MemoryStream();
    
    await workbook.WriteAsync(stream, leaveOpen: true);
    
    // Stream should remain open
    Assert.IsTrue(stream.CanWrite);
    Assert.IsTrue(stream.CanRead);
}

[Test]
public async Task TestCancellation()
{
    using var workbook = CreateLargeWorkbook();
    using var stream = new MemoryStream();
    using var cts = new CancellationTokenSource();
    
    var writeTask = workbook.WriteAsync(stream, cancellationToken: cts.Token);
    
    // Cancel after short delay
    await Task.Delay(100);
    cts.Cancel();
    
    await Assert.ThrowsAsync<OperationCanceledException>(() => writeTask);
}
```

##### Error Handling Tests
```csharp
[Test]
public async Task TestWriteAsyncToClosedStream()
{
    using var workbook = new XSSFWorkbook();
    var stream = new MemoryStream();
    stream.Close();
    
    await Assert.ThrowsAsync<ObjectDisposedException>(() => 
        workbook.WriteAsync(stream));
}

[Test]
public async Task TestWriteAsyncToReadOnlyStream()
{
    using var workbook = new XSSFWorkbook();
    var readOnlyStream = new MemoryStream(new byte[1000], false);
    
    await Assert.ThrowsAsync<NotSupportedException>(() => 
        workbook.WriteAsync(readOnlyStream));
}
```

##### Resource Management Tests
```csharp
[Test]
public async Task TestNoMemoryLeaksAsync()
{
    long initialMemory = GC.GetTotalMemory(true);
    
    for (int i = 0; i < 100; i++)
    {
        using var workbook = CreateComplexWorkbook();
        using var stream = new MemoryStream();
        await workbook.WriteAsync(stream);
    }
    
    GC.Collect();
    GC.WaitForPendingFinalizers();
    GC.Collect();
    
    long finalMemory = GC.GetTotalMemory(true);
    long memoryDifference = finalMemory - initialMemory;
    
    // Allow for some variance but catch major leaks
    Assert.Less(memoryDifference, 50 * 1024 * 1024); // 50MB threshold
}
```

#### 2. Integration Tests

##### Complex Workbook Tests
**Location:** `/testcases/ooxml/XSSF/UserModel/TestXSSFWorkbookAsyncIntegration.cs`

```csharp
[Test]
public async Task TestComplexWorkbookAsync()
{
    using var workbook = new XSSFWorkbook();
    
    // Create complex workbook with multiple sheets, styles, formulas, images
    var sheet1 = CreateWorksheetWithFormulas(workbook, "Formulas");
    var sheet2 = CreateWorksheetWithImages(workbook, "Images");
    var sheet3 = CreateWorksheetWithCharts(workbook, "Charts");
    
    using var stream = new MemoryStream();
    await workbook.WriteAsync(stream);
    
    // Verify by reading back
    stream.Position = 0;
    using var readWorkbook = new XSSFWorkbook(stream);
    
    Assert.AreEqual(3, readWorkbook.NumberOfSheets);
    // Additional verification...
}

[Test]
public async Task TestLargeDatasetAsync()
{
    using var workbook = new XSSFWorkbook();
    var sheet = workbook.CreateSheet("LargeData");
    
    // Create 100k rows with data
    for (int i = 0; i < 100000; i++)
    {
        var row = sheet.CreateRow(i);
        for (int j = 0; j < 10; j++)
        {
            row.CreateCell(j).SetCellValue($"Cell_{i}_{j}");
        }
    }
    
    using var stream = new MemoryStream();
    var stopwatch = Stopwatch.StartNew();
    
    await workbook.WriteAsync(stream);
    
    stopwatch.Stop();
    Assert.Greater(stream.Length, 1000000); // At least 1MB
    Assert.Less(stopwatch.ElapsedMilliseconds, 30000); // Less than 30 seconds
}
```

##### Concurrent Operations Tests
```csharp
[Test]
public async Task TestConcurrentWriteAsync()
{
    const int concurrentOperations = 10;
    var tasks = new Task[concurrentOperations];
    
    for (int i = 0; i < concurrentOperations; i++)
    {
        int taskIndex = i;
        tasks[i] = Task.Run(async () =>
        {
            using var workbook = CreateTestWorkbook(taskIndex);
            using var stream = new MemoryStream();
            await workbook.WriteAsync(stream);
            return stream.ToArray();
        });
    }
    
    var results = await Task.WhenAll(tasks);
    
    // Verify all operations completed successfully
    foreach (var result in results)
    {
        Assert.Greater(result.Length, 1000);
    }
}
```

#### 3. Performance Tests

##### Benchmark Tests
**Location:** `/benchmarks/NPOI.Benchmarks/AsyncWriteBenchmark.cs`

```csharp
[MemoryDiagnoser]
[SimpleJob(RuntimeMoniker.Net60)]
public class AsyncWriteBenchmark
{
    private XSSFWorkbook _workbook;
    private byte[] _syncResult;
    
    [GlobalSetup]
    public void Setup()
    {
        _workbook = CreateBenchmarkWorkbook();
        
        // Get sync baseline
        using var syncStream = new MemoryStream();
        _workbook.Write(syncStream);
        _syncResult = syncStream.ToArray();
    }
    
    [Benchmark(Baseline = true)]
    public byte[] WritSync()
    {
        using var stream = new MemoryStream();
        _workbook.Write(stream);
        return stream.ToArray();
    }
    
    [Benchmark]
    public async Task<byte[]> WriteAsync()
    {
        using var stream = new MemoryStream();
        await _workbook.WriteAsync(stream);
        return stream.ToArray();
    }
    
    [Benchmark]
    public async Task<byte[]> WriteAsyncWithProgress()
    {
        using var stream = new MemoryStream();
        var progress = new Progress<WriteProgress>();
        await _workbook.WriteAsync(stream, progress: progress);
        return stream.ToArray();
    }
}
```

##### Memory Usage Tests
```csharp
[Test]
public async Task TestMemoryUsageAsync()
{
    var memoryBefore = GC.GetTotalMemory(true);
    
    using var workbook = CreateLargeWorkbook(50000); // 50k rows
    using var stream = new MemoryStream();
    
    var memoryDuringCreation = GC.GetTotalMemory(false);
    
    await workbook.WriteAsync(stream);
    
    var memoryAfterWrite = GC.GetTotalMemory(false);
    var memoryIncrease = memoryAfterWrite - memoryBefore;
    
    // Memory usage should be reasonable
    Assert.Less(memoryIncrease, 500 * 1024 * 1024); // 500MB max
    
    // Verify result size
    Assert.Greater(stream.Length, 10 * 1024 * 1024); // At least 10MB
}
```

### Test Data Generation

#### Helper Methods
```csharp
private static XSSFWorkbook CreateTestWorkbook(int seed = 0)
{
    var workbook = new XSSFWorkbook();
    var sheet = workbook.CreateSheet($"Sheet{seed}");
    
    // Add various content types
    for (int i = 0; i < 100; i++)
    {
        var row = sheet.CreateRow(i);
        row.CreateCell(0).SetCellValue($"String {seed}_{i}");
        row.CreateCell(1).SetCellValue(DateTime.Now.AddDays(i));
        row.CreateCell(2).SetCellValue(i * seed + 1.5);
        row.CreateCell(3).SetCellFormula($"B{i + 1}+C{i + 1}");
    }
    
    return workbook;
}

private static XSSFWorkbook CreateLargeWorkbook(int rowCount = 10000)
{
    var workbook = new XSSFWorkbook();
    var sheet = workbook.CreateSheet("LargeSheet");
    
    for (int i = 0; i < rowCount; i++)
    {
        var row = sheet.CreateRow(i);
        for (int j = 0; j < 20; j++)
        {
            row.CreateCell(j).SetCellValue($"Cell_{i}_{j}");
        }
    }
    
    return workbook;
}
```

### Success Criteria

#### Functional Correctness
- All async operations produce identical results to sync operations
- Proper exception handling and error propagation
- Reliable cancellation support
- Resource cleanup verification

#### Performance Requirements
- Async operations should not be slower than sync operations
- Memory usage should not exceed sync operations by more than 20%
- Large workbook operations should show measurable async benefits
- Scalability under concurrent operations

#### Reliability Standards
- Zero memory leaks in sustained testing
- Proper disposal of all resources
- Thread-safe operation where applicable
- Graceful degradation under resource constraints

---

## Implementation Notes

### General Guidelines
- Use `ConfigureAwait(false)` consistently throughout
- Maintain backward compatibility with all existing APIs
- Follow existing NPOI coding conventions and patterns
- Focus on establishing async pipeline in Phase 1 before optimizing
- Each phase should be fully functional before moving to the next

### Performance Considerations
- Phase 1 may use some async-over-sync patterns that will be optimized later
- Phase 2 should show significant async performance benefits
- Phase 3 should demonstrate production-ready optimization
- Memory usage and resource management are critical throughout

### Testing Approach
- Run existing synchronous tests alongside new async tests
- Ensure bit-identical output between sync and async versions
- Performance regression testing at each phase
- Comprehensive cancellation and error handling validation