using System.ComponentModel.DataAnnotations;
using System.Reflection;
using BcWordLayout.Domain;
using BcWordLayout.McpHost.Tools;
using ModelContextProtocol.Server;

namespace BcWordLayout.Tests;

/// <summary>
/// The allowed values an enumerated argument publishes in its input schema (<c>AllowedValueAttributes.cs</c>, GitHub
/// issue #133) and the parser that reads the argument must agree: every listed value is one the parser accepts, and
/// every parameter of that name carries its list. Otherwise tools/list would offer a value the server refuses, or a
/// new tool would publish a bare string again.
/// </summary>
public class AllowedValueAttributesTests
{
    [Fact]
    public void Every_listed_location_type_is_accepted() =>
        AssertEachAccepted<LocationTypeValuesAttribute>(value => ToolGuards.BuildLocation(value, null, null, null, null, null));

    [Fact]
    public void Every_listed_layout_part_is_accepted() =>
        AssertEachAccepted<LayoutPartValuesAttribute>(value => Assert.True(ToolGuards.TryParseLayoutPart(value, out _), value));

    [Fact]
    public void Every_listed_alignment_is_accepted() =>
        AssertEachAccepted<AlignmentValuesAttribute>(value => CellTextEditor.ParseAlignment(value, "alignment"));

    [Fact]
    public void Every_listed_borders_value_is_accepted() =>
        AssertEachAccepted<BordersValuesAttribute>(value => TableTools.ParseBordersOrThrow(value, default));

    [Fact]
    public void Every_listed_border_style_is_accepted() =>
        AssertEachAccepted<BorderStyleValuesAttribute>(value => TableTools.ParseCellBorderOptionsOrThrow(["top"], value, null));

    [Fact]
    public void Every_listed_edge_is_accepted() =>
        AssertEachAccepted<EdgeValuesAttribute>(value => TableTools.ParseCellBorderOptionsOrThrow([value], null, null));

    [Fact]
    public void Every_listed_column_mode_is_accepted() =>
        AssertEachAccepted<ColumnModeValuesAttribute>(value => TableTools.ParseInsertColumnModeOrThrow(value));

    [Fact]
    public void Every_listed_text_format_target_is_accepted() =>
        AssertEachAccepted<TargetValuesAttribute>(value => EditTools.ParseTextFormatTargetOrThrow(value));

    [Fact]
    public void Every_listed_validation_level_is_accepted() =>
        AssertEachAccepted<LevelValuesAttribute>(value => Assert.True(ReadTools.TryParseLevel(value, out _), value));

    [Fact]
    public void Every_listed_converter_is_accepted() =>
        AssertEachAccepted<ConverterValuesAttribute>(value => Assert.True(LifecycleTools.TryParseConverterKind(value, out _), value));

    /// <summary>The list each enumerated argument name carries, wherever it appears.</summary>
    private static readonly Dictionary<string, Type> ListByName = new(StringComparer.Ordinal)
    {
        ["locationType"] = typeof(LocationTypeValuesAttribute),
        ["layoutPart"] = typeof(LayoutPartValuesAttribute),
        ["alignment"] = typeof(AlignmentValuesAttribute),
        ["borders"] = typeof(BordersValuesAttribute),
        ["borderStyle"] = typeof(BorderStyleValuesAttribute),
        ["edges"] = typeof(EdgeValuesAttribute),
        ["mode"] = typeof(ColumnModeValuesAttribute),
        ["target"] = typeof(TargetValuesAttribute),
        ["level"] = typeof(LevelValuesAttribute),
        ["converter"] = typeof(ConverterValuesAttribute),
    };

    [Fact]
    public void Every_enumerated_parameter_and_member_publishes_its_values()
    {
        var tools = typeof(ToolGuards).Assembly.GetTypes()
            .Where(t => t.GetCustomAttribute<McpServerToolTypeAttribute>() is not null)
            .SelectMany(t => t.GetMethods(BindingFlags.Public | BindingFlags.Static))
            .Where(m => m.GetCustomAttribute<McpServerToolAttribute>() is not null)
            .ToList();

        var checkedCount = 0;
        var missing = new List<string>();
        foreach (var method in tools)
        {
            var tool = method.GetCustomAttribute<McpServerToolAttribute>()!.Name;
            foreach (var parameter in method.GetParameters())
            {
                Check($"{tool}.{parameter.Name}", parameter.Name!, parameter);
                var element = parameter.ParameterType.IsGenericType ? parameter.ParameterType.GetGenericArguments()[0] : parameter.ParameterType;
                if (element.IsClass && element != typeof(string) && element.Namespace == typeof(TableTools).Namespace)
                {
                    foreach (var property in element.GetProperties())
                    {
                        Check($"{tool}.{parameter.Name}.{property.Name}", char.ToLowerInvariant(property.Name[0]) + property.Name[1..], property);
                    }
                }
            }
        }

        Assert.True(checkedCount >= 30, $"only {checkedCount} enumerated parameters found - has the naming changed?");
        Assert.True(missing.Count == 0, "Enumerated arguments without their allowed values: " + string.Join(", ", missing));

        void Check(string where, string name, ICustomAttributeProvider target)
        {
            if (!ListByName.TryGetValue(name, out var list))
            {
                return;
            }

            checkedCount++;
            if (target.GetCustomAttributes(list, inherit: false).Length == 0)
            {
                missing.Add(where);
            }
        }
    }

    private static void AssertEachAccepted<TList>(Action<string> parse)
        where TList : AllowedValuesAttribute, new()
    {
        var values = new TList().Values.OfType<string>().ToList();
        Assert.NotEmpty(values);
        foreach (var value in values)
        {
            parse(value);
        }
    }
}
