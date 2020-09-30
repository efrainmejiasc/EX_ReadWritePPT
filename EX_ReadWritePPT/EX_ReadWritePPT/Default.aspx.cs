using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;



public partial class _Default : Page
{
    private string pathPresentacion = "~/TemplatePPT/Plantilla1.pptx";

    protected void Page_Load(object sender, EventArgs e)
    {

        EscribirUsandoOffice15();
    }



    private void EscribirUsandoOffice15()
    {
        Label2.Text = "Verifique el archivo: " + System.Web.HttpContext.Current.Server.MapPath(this.pathPresentacion);

        ReadWriteSlide write = new ReadWriteSlide();
        if (write.WriteOnSlide())
            Label1.Text = "Se escribio correctamente en la presentacion";
        else
            Label1.Text = "Error escribiendo  en la presentacion";
    }
}