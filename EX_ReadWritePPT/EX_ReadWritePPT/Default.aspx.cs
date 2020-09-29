using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;



public partial class _Default : Page
{

    protected void Page_Load(object sender, EventArgs e)
    {
        Label2.Text = "Verifique el archivo: "  + System.Web.HttpContext.Current.Server.MapPath("~/TemplatePPT/Plantilla1.pptx");

        ReadWriteSlide write = new ReadWriteSlide();
        if (write.WriteOnSlide())
            Label1.Text = "Se escribio correctamente en la presentacion";
        else
            Label1.Text = "Error escribiendo  en la presentacion";
    }
}