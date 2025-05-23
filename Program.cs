using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml;

class Program
{
    static void AbrirArchivo(string ruta)
    {
        try
        {
            using var p = new Process();
            p.StartInfo = new ProcessStartInfo(ruta) { UseShellExecute = true };
            p.Start();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[WARN] No se pudo abrir {Path.GetFileName(ruta)}: {ex.Message}");
        }
    }

    static async Task Main(string[] args)
    {
        Console.WriteLine("Tema de investigación:");
        string prompt = Console.ReadLine();

        string resultado = await ConsultarAIAsync(prompt);

        Console.WriteLine("\nResultado de la investigación:\n");
        Console.WriteLine(resultado);
        Console.WriteLine("\n¿Desea modificar el resultado antes de guardar? (s/n)");
        if (Console.ReadLine().Trim().ToLower() == "s")
        {
            Console.WriteLine("Ingrese el nuevo resultado:");
            resultado = Console.ReadLine();
        }

        GuardarInvestigacion(prompt, resultado);

        string carpeta = Path.Combine(Environment.CurrentDirectory, "Informe_" + DateTime.Now.ToString("yyyyMMdd_HHmmss"));
        Directory.CreateDirectory(carpeta);
        Console.WriteLine("[DEBUG] Contenido a guardar:\n" + resultado);

        string rutaWord = Path.Combine(carpeta, "Informe.docx");
        string rutaPptx = Path.Combine(carpeta, "Presentacion.pptx");

        GenerarWord(rutaWord, resultado);
        GenerarPowerPoint(rutaPptx, resultado);

        AbrirArchivo(rutaWord);
        AbrirArchivo(rutaPptx);

        Console.WriteLine($"\nArchivos generados en: {carpeta}");
    }

    static async Task<string> ConsultarAIAsync(string prompt)
    {
        var apiKey = "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"; // ⚠️ Asegúrate de mantener tu API key segura
        var url = "https://api.openai.com/v1/chat/completions";

        var requestBody = new
        {
            model = "gpt-3.5-turbo",
            messages = new[] { new { role = "user", content = prompt } }
        };

        using var client = new HttpClient();
        client.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}");

        var content = new StringContent(JsonSerializer.Serialize(requestBody), Encoding.UTF8, "application/json");
        var response = await client.PostAsync(url, content);
        var responseString = await response.Content.ReadAsStringAsync();

        try
        {
            using var doc = JsonDocument.Parse(responseString);
            string result = doc.RootElement
                .GetProperty("choices")[0]
                .GetProperty("message")
                .GetProperty("content")
                .GetString();

            return result.Trim();
        }
        catch (Exception ex)
        {
            return $"Error: No se recibió una respuesta válida de la API.\nDetalles: {ex.Message}\n\n{responseString}";
        }
    }

    static void GuardarInvestigacion(string prompt, string resultado)
    {
        Console.WriteLine("Guardando en base de datos...");
        // Lógica futura aquí
    }

    static void GenerarWord(string ruta, string contenido)
    {
        using var doc = WordprocessingDocument.Create(ruta, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document();
        var body = new Body(new Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text(contenido))));
        mainPart.Document.Append(body);
        mainPart.Document.Save();
    }

    static void GenerarPowerPoint(string ruta, string contenido)
    {
        using var presentationDoc = PresentationDocument.Create(ruta, PresentationDocumentType.Presentation);

        var presentationPart = presentationDoc.AddPresentationPart();
        presentationPart.Presentation = new Presentation();

        var slidePart = presentationPart.AddNewPart<SlidePart>();
        slidePart.Slide = new Slide(
            new CommonSlideData(
                new ShapeTree(
                    new NonVisualGroupShapeProperties(
                        new NonVisualDrawingProperties() { Id = 1U, Name = "Title Slide" },
                        new NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()
                    ),
                    new GroupShapeProperties(),
                    new Shape(
                        new NonVisualShapeProperties(
                            new NonVisualDrawingProperties() { Id = 2U, Name = "Title" },
                            new NonVisualShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()
                        ),
                        new ShapeProperties(),
                        new TextBody(
                            new A.BodyProperties(),
                            new A.ListStyle(),
                            new A.Paragraph(
                                new A.Run(
                                    new A.RunProperties() { Language = "es-ES", FontSize = 2400 },
                                    new A.Text(contenido)
                                )
                            )
                        )
                    )
                )
            )
        );
        slidePart.Slide.Save();

        var slideIdList = new SlideIdList();
        slideIdList.Append(new SlideId() { Id = 256U, RelationshipId = presentationPart.GetIdOfPart(slidePart) });

        presentationPart.Presentation.Append(slideIdList);
        presentationPart.Presentation.Save();
    }
}