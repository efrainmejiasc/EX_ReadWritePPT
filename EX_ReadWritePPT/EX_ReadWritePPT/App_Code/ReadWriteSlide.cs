using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

/// <summary>
/// Descripción breve de ReadWriteSlide
/// </summary>
public class ReadWriteSlide
{
    public ReadWriteSlide()
    {
    }

    public bool WriteOnSlide()
    {
        var resultado = false;
        string filePath = System.Web.HttpContext.Current.Server.MapPath("~/TemplatePPT/Plantilla1.pptx");
        try
        {
            Application pptApplication = new Application();
            Presentations multi_presentations = pptApplication.Presentations;
            Presentation presentation = multi_presentations.Open(filePath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
            CustomLayout customLayout = presentation.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutText];

            Slides slides = presentation.Slides;
            Microsoft.Office.Interop.PowerPoint.Shapes shapes = presentation.Slides[1].Shapes;
            TextRange objText;

            slides = presentation.Slides;

            var text1 = "Escribiendo en el titulo: " + DateTime.Now.ToString();
            var text2 = "Descripcion PPT: " + DateTime.Now.ToString();

            objText = shapes[1].TextFrame.TextRange;
            objText.Text = text1;
            objText.Font.Name = "Arial";
            objText.Font.Size = 32;

            objText = shapes[2].TextFrame.TextRange;
            objText.Text = text2;
            objText.Font.Name = "Arial";
            objText.Font.Size = 28;

            ReadWriteTxt(filePath);
            presentation.SaveAs(filePath, PpSaveAsFileType.ppSaveAsDefault, Microsoft.Office.Core.MsoTriState.msoTrue);
            presentation.Close();
            pptApplication.Quit();
            resultado = true;
        }
        catch (Exception ex)
        {

        }
        return resultado;
    }

    public string  ReadSlide()
    {
        string presentationText = string.Empty;
        try
        {
            string filePath = System.Web.HttpContext.Current.Server.MapPath("~/TemplatePPT/Plantilla1.pptx");

            Application pptApplication = new Application();
            Presentations multi_presentations = pptApplication.Presentations;
            Presentation presentation = multi_presentations.Open(filePath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
            foreach (var item in presentation.Slides[1].Shapes)
            {
                var shape = (Microsoft.Office.Interop.PowerPoint.Shape)item;
                if (shape.HasTextFrame == MsoTriState.msoTrue)
                {
                    if (shape.TextFrame.HasText == MsoTriState.msoTrue)
                    {
                        var textRange = shape.TextFrame.TextRange;
                        var text = textRange.Text;

                        presentationText += text + " ";
                    }
                }
            }

            Console.WriteLine(presentationText);
        }
        catch (Exception ex)
        {
          
        }
        return presentationText;
    }

    private void ReadWriteTxt(string pathArchivo)
    {
        FileAttributes atr = File.GetAttributes(pathArchivo);
        File.SetAttributes(pathArchivo, atr & ~FileAttributes.ReadOnly);
    }

}