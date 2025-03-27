using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using DocumentConverterApplication.Services;
using Microsoft.AspNetCore.Http.Features;
using System.IO;

namespace DocumentConverterApplication
{
    public class Startup
    {
        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        public IConfiguration Configuration { get; }

        public void ConfigureServices(IServiceCollection services)
        {
            // Add support for MVC
            services.AddControllersWithViews();

            // Configure detailed file upload limits
            services.Configure<FormOptions>(options =>
            {
                options.MultipartBodyLengthLimit = 20 * 1024 * 1024; // 20MB total file size
                options.MemoryBufferThreshold = 2 * 1024 * 1024; // 2MB memory buffer
                options.ValueLengthLimit = int.MaxValue; // Maximum value length
                options.MultipartBoundaryLengthLimit = int.MaxValue; // Maximum boundary length
                options.MultipartHeadersCountLimit = int.MaxValue; // Maximum headers count
                options.MultipartHeadersLengthLimit = int.MaxValue; // Maximum headers length
            });

            // Register services
            services.AddScoped<ITempFileService, TempFileService>();
            services.AddLogging(); // Add logging
        }

        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }
            else
            {
                app.UseExceptionHandler("/Home/Error");
                app.UseHsts();
            }

            app.UseHttpsRedirection();
            app.UseStaticFiles();

            // Ensure downloads directory exists
            var downloadsPath = Path.Combine(env.WebRootPath, "downloads");
            Directory.CreateDirectory(downloadsPath);

            app.UseRouting();
            app.UseAuthorization();

            app.UseEndpoints(endpoints =>
            {
                endpoints.MapControllerRoute(
                    name: "default",
                    pattern: "{controller=Home}/{action=Index}/{id?}");
            });
        }
    }
}