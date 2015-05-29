using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(Gps.Web.Startup))]
namespace Gps.Web
{
    public partial class Startup {
        public void Configuration(IAppBuilder app) {
            ConfigureAuth(app);
        }
    }
}
