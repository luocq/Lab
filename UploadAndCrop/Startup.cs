using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(UploadAndCrop.Startup))]
namespace UploadAndCrop
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
