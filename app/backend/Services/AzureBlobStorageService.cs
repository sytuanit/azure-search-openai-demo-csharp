// Copyright (c) Microsoft. All rights reserved.

using System.Threading;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
using static System.Net.Mime.MediaTypeNames;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace MinimalApi.Services;

internal sealed class AzureBlobStorageService(BlobContainerClient container)
{
    private static List<string> s_wordDocumentExtensions = new List<string>()
    {
        ".doc", ".docx","dotx"
    };
    private static List<string> s_excelExtensions = new List<string>()
    {
        ".xlsm",".xlsb",".xls",".xlsx"
    };
    private static List<string> s_powerPointExtensions = new List<string>()
    {
        ".ppt", ".pptx"
    };
    private static List<string> s_csvExtensions = new List<string>()
    {
        ".csv"
    };
    internal static DefaultAzureCredential DefaultCredential { get; } = new();

    internal async Task<UploadDocumentsResponse> UploadFilesAsync(IEnumerable<IFormFile> files, CancellationToken cancellationToken)
    {
        try
        {
            List<string> uploadedFiles = [];
            foreach (var file in files)
            {
                var fileName = file.FileName;
                if (s_wordDocumentExtensions.Any(n => n.Equals(Path.GetExtension(fileName).ToLower())))
                {
                    uploadedFiles.AddRange(await InternalUploadWordFileAsync(file, cancellationToken));
                }
                else if (s_excelExtensions.Any(n => n.Equals(Path.GetExtension(fileName).ToLower())))
                {
                    uploadedFiles.AddRange(await InternalUploadExcelFileAsync(file, cancellationToken));
                }
                else if (s_powerPointExtensions.Any(n => n.Equals(Path.GetExtension(fileName).ToLower())))
                {
                    uploadedFiles.AddRange(await InternalUploadPowerPointFileAsync(file, cancellationToken));
                }
                else if (s_csvExtensions.Any(n => n.Equals(Path.GetExtension(fileName).ToLower())))
                {
                    uploadedFiles.AddRange(await InternalUploadExcelFileAsync(file, cancellationToken));
                }
                else if (Path.GetExtension(fileName).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
                {
                    await using var stream = file.OpenReadStream();
#pragma warning disable CA2000 // Dispose objects before losing scope
                    using var documents = PdfReader.Open(stream, PdfDocumentOpenMode.Import);
#pragma warning restore CA2000 // Dispose objects before losing scope
                    uploadedFiles.AddRange(await InternalUploadPdfFileAsync(fileName, documents, cancellationToken));
                }
            }

            if (uploadedFiles.Count is 0)
            {
                return UploadDocumentsResponse.FromError("""
                    No files were uploaded. Either the files already exist or the file format are not supported.
                    """);
            }

            return new UploadDocumentsResponse([.. uploadedFiles]);
        }
#pragma warning disable CA1031 // Do not catch general exception types
        catch (Exception ex)
        {
            return UploadDocumentsResponse.FromError(ex.ToString());
        }
#pragma warning restore CA1031 // Do not catch general exception types
    }

    private byte[] ReadToEnd(Stream stream)
    {
        long originalPosition = 0;

        if (stream.CanSeek)
        {
            originalPosition = stream.Position;
            stream.Position = 0;
        }

        try
        {
            byte[] readBuffer = new byte[4096];

            int totalBytesRead = 0;
            int bytesRead;

            while ((bytesRead = stream.Read(readBuffer, totalBytesRead, readBuffer.Length - totalBytesRead)) > 0)
            {
                totalBytesRead += bytesRead;

                if (totalBytesRead == readBuffer.Length)
                {
                    int nextByte = stream.ReadByte();
                    if (nextByte != -1)
                    {
                        byte[] temp = new byte[readBuffer.Length * 2];
                        Buffer.BlockCopy(readBuffer, 0, temp, 0, readBuffer.Length);
                        Buffer.SetByte(temp, totalBytesRead, (byte)nextByte);
                        readBuffer = temp;
                        totalBytesRead++;
                    }
                }
            }

            byte[] buffer = readBuffer;
            if (readBuffer.Length != totalBytesRead)
            {
                buffer = new byte[totalBytesRead];
                Buffer.BlockCopy(readBuffer, 0, buffer, 0, totalBytesRead);
            }

            return buffer;
        }
        finally
        {
            if (stream.CanSeek)
            {
                stream.Position = originalPosition;
            }
        }
    }

    private async Task<List<string>> InternalUploadWordFileAsync(IFormFile file, CancellationToken cancellationToken)
    {
        var fileName = file.FileName;
        await using var fileStreamInput = file.OpenReadStream();

        //Loads file stream into Word document
        using WordDocument wordDocument = new WordDocument(fileStreamInput, Syncfusion.DocIO.FormatType.Automatic);

        //Instantiation of DocIORenderer for Word to PDF conversion
        using DocIORenderer render = new DocIORenderer();

        //Converts Word document into PDF document
        Syncfusion.Pdf.PdfDocument pdfDocument = render.ConvertToPDF(wordDocument);

        //Saves the PDF document to MemoryStream.
        MemoryStream stream = new MemoryStream();
        pdfDocument.Save(stream);
        stream.Position = 0;

        fileName = Path.GetFileNameWithoutExtension(fileName) + ".pdf";
        using var documents = PdfReader.Open(stream, PdfDocumentOpenMode.Import);
        return await InternalUploadPdfFileAsync(fileName, documents, cancellationToken);
    }

    private async Task<List<string>> InternalUploadExcelFileAsync(IFormFile file, CancellationToken cancellationToken)
    {
        using ExcelEngine excelEngine = new ExcelEngine();
        IApplication application = excelEngine.Excel;
        application.DefaultVersion = ExcelVersion.Xlsx;

        var fileName = file.FileName;
        await using var fileStreamInput = file.OpenReadStream();
        IWorkbook workbook = application.Workbooks.Open(fileStreamInput);

        //Initialize XlsIO renderer.
        XlsIORenderer renderer = new XlsIORenderer();

        //Convert Excel document into PDF document 
        Syncfusion.Pdf.PdfDocument pdfDocument = renderer.ConvertToPDF(workbook);

        //Saves the PDF document to MemoryStream.
        MemoryStream stream = new MemoryStream();
        pdfDocument.Save(stream);
        stream.Position = 0;

        fileName = Path.GetFileNameWithoutExtension(fileName) + ".pdf";
        using var documents = PdfReader.Open(stream, PdfDocumentOpenMode.Import);
        return await InternalUploadPdfFileAsync(fileName, documents, cancellationToken);
    }

    private async Task<List<string>> InternalUploadPowerPointFileAsync(IFormFile file, CancellationToken cancellationToken)
    {
        var fileName = file.FileName;
        await using var fileStreamInput = file.OpenReadStream();

        //Open the existing PowerPoint presentation with loaded stream.
        using IPresentation pptxDoc = Presentation.Open(fileStreamInput);

        //Create the MemoryStream to save the converted PDF.
        using MemoryStream pdfStream = new MemoryStream();

        //Convert the PowerPoint document to PDF document.
        using Syncfusion.Pdf.PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc);

        //Saves the PDF document to MemoryStream.
        MemoryStream stream = new MemoryStream();
        pdfDocument.Save(stream);
        stream.Position = 0;

        fileName = Path.GetFileNameWithoutExtension(fileName) + ".pdf";
        using var documents = PdfReader.Open(stream, PdfDocumentOpenMode.Import);
        return await InternalUploadPdfFileAsync(fileName, documents, cancellationToken);
    }

    private async Task<List<string>> InternalUploadPdfFileAsync(string fileName, PdfDocument documents, CancellationToken cancellationToken)
    {
        List<string> uploadedFiles = [];
        for (int i = 0; i < documents.PageCount; i++)
        {
            var documentName = BlobNameFromFilePage(fileName, i);
            var blobClient = container.GetBlobClient(documentName);
            if (await blobClient.ExistsAsync(cancellationToken))
            {
                continue;
            }

            var tempFileName = Path.GetTempFileName();

            try
            {
                using var document = new PdfDocument();
                document.AddPage(documents.Pages[i]);
                document.Save(tempFileName);

                await using var tempStream = File.OpenRead(tempFileName);
                await blobClient.UploadAsync(tempStream, new BlobHttpHeaders
                {
                    ContentType = "application/pdf"
                }, cancellationToken: cancellationToken);

                uploadedFiles.Add(documentName);
            }
            finally
            {
                File.Delete(tempFileName);
            }
        }
        return uploadedFiles;
    }

    private static string BlobNameFromFilePage(string filename, int page = 0) =>
        Path.GetExtension(filename).ToLower() is ".pdf"
            ? $"{Path.GetFileNameWithoutExtension(filename)}-{page}.pdf"
            : Path.GetFileName(filename);
}
