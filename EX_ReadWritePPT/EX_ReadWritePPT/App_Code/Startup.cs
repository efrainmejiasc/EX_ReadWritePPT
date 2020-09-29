using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(EX_ReadWritePPT.Startup))]
namespace EX_ReadWritePPT
{
    public partial class Startup {
        public void Configuration(IAppBuilder app) {
            ConfigureAuth(app);
        }
    }
}
