using Microsoft.Extensions.DependencyInjection;
using PptxXML.Services;

namespace PptxXML.Utilities
{
    /// <summary>
    /// Contains extension methods for the <see cref="IServiceCollection"/> interface.
    /// </summary>
    public static class IServiceCollectionExtensions
    {
        /// <summary>
        /// Register the PptxXML library's dependencies.
        /// </summary>
        public static IServiceCollection AddPptxXMLLibrary(this IServiceCollection services)
        {
            services.AddScoped<IElementFactory, ElementFactory>();
            services.AddScoped<IGroupShapeTypeParser, GroupShapeTypeParser>();
            services.AddScoped<IGroupShapeTypeParser, GroupShapeTypeParser>();

            return services;
        }
    }
}
