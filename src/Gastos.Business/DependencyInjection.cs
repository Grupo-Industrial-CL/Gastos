using Gastos.Business.Proveedores;
using Gastos.Domain.Abstractions;
using Gastos.Domain.Abstractions.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace Gastos.Business;

public static class DependencyInjection
{
    /// <summary>
    /// Registra los servicios de negocio por convención (IXxxService -> XxxService, Scoped) y los clientes de las
    /// APIs de Repsol y Moeve (secciones de configuración Proveedores:Repsol y Proveedores:Moeve).
    /// La capa de datos (AddGastosData) la registra el host, que conoce la cadena de conexión.
    /// </summary>
    public static IServiceCollection AddGastosBusiness(this IServiceCollection services, IConfiguration configuration)
    {
        var business = typeof(DependencyInjection).Assembly;
        var domain = typeof(IGastoService).Assembly;

        var interfaces = domain.GetTypes().Where(t => t.IsInterface && t.Namespace == "Gastos.Domain.Abstractions.Services" && t.Name.StartsWith('I'));
        var impls = business.GetTypes().Where(t => t is { IsClass: true, IsAbstract: false }).ToList();
        foreach (var iface in interfaces)
        {
            var impl = impls.FirstOrDefault(t => t.Name == iface.Name[1..] && iface.IsAssignableFrom(t))
                       ?? impls.FirstOrDefault(iface.IsAssignableFrom);
            if (impl is not null) services.AddScoped(iface, impl);
        }

        services.Configure<ProveedorApiOptions>(RepsolApiClient.Nombre, configuration.GetSection("Proveedores:Repsol"));
        services.Configure<ProveedorApiOptions>(MoeveApiClient.Nombre, configuration.GetSection("Proveedores:Moeve"));
        services.AddHttpClient<RepsolApiClient>();
        services.AddHttpClient<MoeveApiClient>();
        services.AddScoped<IProveedorTarjetaCarburante>(sp => sp.GetRequiredService<RepsolApiClient>());
        services.AddScoped<IProveedorTarjetaCarburante>(sp => sp.GetRequiredService<MoeveApiClient>());
        return services;
    }
}
