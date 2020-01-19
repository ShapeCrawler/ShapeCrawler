using Microsoft.Extensions.DependencyInjection;
using SlideXML.Services;

namespace SlideXML.Utilities
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
            services.AddTransient<IElementFactory, ElementFactory>();
            services.AddTransient<IGroupShapeTypeParser, GroupShapeTypeParser>();

            return services;
        }
    }
}
