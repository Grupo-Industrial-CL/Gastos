using System.Collections.Concurrent;
using System.ComponentModel.DataAnnotations.Schema;
using System.Reflection;
using Dapper;

namespace Gastos.Data;

/// <summary>
/// Hace que Dapper respete [Column("...")] en las entidades del dominio (p. ej. Id_Epsilon -> IdEpsilon),
/// cayendo al nombre de la propiedad cuando no hay alias. Copiado del planteamiento de MasterPlan v2.
/// </summary>
public static class DapperColumnMapping
{
    private static bool _initialized;
    private static readonly object _lock = new();

    public static void Initialize(params Assembly[] assemblies)
    {
        if (_initialized) return;
        lock (_lock)
        {
            if (_initialized) return;
            foreach (var asm in assemblies)
                foreach (var type in SafeGetTypes(asm))
                {
                    if (!type.IsClass || type.IsAbstract) continue;
                    if (type.GetProperties().Any(p => p.GetCustomAttribute<ColumnAttribute>() is not null))
                        SqlMapper.SetTypeMap(type, new ColumnAttributeTypeMapper(type));
                }
            _initialized = true;
        }
    }

    private static IEnumerable<Type> SafeGetTypes(Assembly asm)
    {
        try { return asm.GetTypes(); }
        catch (ReflectionTypeLoadException ex) { return ex.Types.Where(t => t is not null)!; }
    }

    private sealed class ColumnAttributeTypeMapper : SqlMapper.ITypeMap
    {
        private readonly Type _type;
        private readonly ConcurrentDictionary<string, PropertyInfo?> _cache = new(StringComparer.OrdinalIgnoreCase);
        public ColumnAttributeTypeMapper(Type type) => _type = type;

        public ConstructorInfo? FindConstructor(string[] names, Type[] types)
            => _type.GetConstructor(Type.EmptyTypes) ?? _type.GetConstructors().FirstOrDefault();
        public ConstructorInfo? FindExplicitConstructor() => null;
        public SqlMapper.IMemberMap? GetConstructorParameter(ConstructorInfo constructor, string columnName) => null;

        public SqlMapper.IMemberMap? GetMember(string columnName)
        {
            var prop = _cache.GetOrAdd(columnName, name =>
            {
                var props = _type.GetProperties();
                return props.FirstOrDefault(p => string.Equals(p.GetCustomAttribute<ColumnAttribute>()?.Name, name, StringComparison.OrdinalIgnoreCase))
                       ?? props.FirstOrDefault(p => string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase));
            });
            return prop is null ? null : new SimpleMemberMap(columnName, prop);
        }
    }

    private sealed class SimpleMemberMap : SqlMapper.IMemberMap
    {
        public SimpleMemberMap(string columnName, PropertyInfo property) { ColumnName = columnName; Property = property; }
        public string ColumnName { get; }
        public Type MemberType => Property.PropertyType;
        public PropertyInfo Property { get; }
        public FieldInfo? Field => null;
        public ParameterInfo? Parameter => null;
    }
}
