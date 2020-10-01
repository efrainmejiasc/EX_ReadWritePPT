using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

/// <summary>
/// Descripción breve de OpenXmlFormat
/// </summary>
public class OpenXmlDocumentFormat
{
    public OpenXmlDocumentFormat()
    {
    }

    public bool WriteOnSlide()
    {
        string filePath = System.Web.HttpContext.Current.Server.MapPath("~/TemplatePPT/Plantilla1.pptx");
        try
        {
            using (PresentationDocument presentationDocument = PresentationDocument.Open(filePath, true))
            {
                PresentationPart presentationPart = presentationDocument.PresentationPart;
                Presentation presentation = presentationPart.Presentation;
            }
        }
        catch (Exception ex)
        {
            return false;
        }

        return true;
    }
}