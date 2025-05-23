using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace Proyecto1Final
{
    public static class GeneradorPowerPoint
    {
        public static void GenerarPowerPoint(string ruta, string contenido)
        {
            using (var presentationDocument = PresentationDocument.Create(ruta, PresentationDocumentType.Presentation))
            {
                // Crear la parte de la presentación
                var presentationPart = presentationDocument.AddPresentationPart();
                presentationPart.Presentation = new Presentation();

                // Crear una diapositiva
                var slidePart = presentationPart.AddNewPart<SlidePart>();
                slidePart.Slide = new Slide(
                    new CommonSlideData(
                        new ShapeTree(
                            new NonVisualGroupShapeProperties(
                                new NonVisualDrawingProperties() { Id = 1U, Name = "Grupo" },
                                new NonVisualGroupShapeDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()
                            ),
                            new GroupShapeProperties(),
                            new Shape(
                                new NonVisualShapeProperties(
                                    new NonVisualDrawingProperties() { Id = 2U, Name = "TextoContenido" },
                                    new NonVisualShapeDrawingProperties(),
                                    new ApplicationNonVisualDrawingProperties()
                                ),
                                new ShapeProperties(),
                                new TextBody(
                                    new A.BodyProperties(),
                                    new A.ListStyle(),
                                    new A.Paragraph(
                                        new A.Run(
                                            new A.RunProperties() { FontSize = 2400 },
                                            new A.Text(contenido)
                                        )
                                    )
                                )
                            )
                        )
                    )
                );
                slidePart.Slide.Save();

                // Asociar la diapositiva a la presentación
                var slideIdList = new SlideIdList();
                slideIdList.Append(new SlideId()
                {
                    Id = 256U,
                    RelationshipId = presentationPart.GetIdOfPart(slidePart)
                });

                presentationPart.Presentation.Append(slideIdList);
                presentationPart.Presentation.Save();
            }
        }
    }
}
