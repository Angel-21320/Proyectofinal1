static void GenerarWord(string ruta, string contenido)
{
    using var doc = WordprocessingDocument.Create(ruta, WordprocessingDocumentType.Document);
    var mainPart = doc.AddMainDocumentPart();
    mainPart.Document = new Document(
        new Body(
            new Paragraph(
                new Run(
                    new Text(contenido)
                )
            )
        )
    );
    mainPart.Document.Save();
}