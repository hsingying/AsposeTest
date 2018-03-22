using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(ILHG_TEST.Startup))]
namespace ILHG_TEST
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
