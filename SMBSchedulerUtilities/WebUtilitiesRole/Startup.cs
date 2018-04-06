using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(WebUtilitiesRole.Startup))]
namespace WebUtilitiesRole
{
    public partial class Startup {
        public void Configuration(IAppBuilder app) {
            ConfigureAuth(app);
        }
    }
}
