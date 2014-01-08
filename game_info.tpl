{% autoescape None %}
using System.Collections.Generic;
namespace Config
{

{% for k, lst in root.items() %}

{% set struct, unit = "_unit_of_" + k, lst[0] %}

public struct {{ struct }} {
    {% for name, value in unit.items() %}
    public {{ value_type_conv(i[name] for i in lst) }} {{ name }};
    {% end %}
    public {{ struct }}({{ ", ".join("{} _{}".format(value_type_conv(i[k] for i in lst), k) for k, v in unit.items()) }}) {
        {% for name in unit %}
        {{ name }} = _{{ name }};
        {% end %}
    }
}
{% end %}


public class GameConfig {
    {% for k, lst in root.items() %}
    {% set struct, unit, idx_type = "_unit_of_" + k, lst[0], value_type_conv(i[idx_key_map.get(k, "id")] for i in lst) %}
    public static Dictionary<{{ idx_type }}, {{ struct }}> {{ k }} = new Dictionary<{{ idx_type }}, {{ struct }}>();
    {% end %}

    static GameConfig() {
        {% for k, lst in root.items() %}
        {% set struct, key = "_unit_of_" + k, idx_key_map.get(k, "id") %}
        {% for unit in lst %}
        {{ k }}.Add({{ json_encode(unit[key]) }}, new {{ struct }}({{ ", ".join(value_conv(v) for v in unit.values()) }}));
        {% end %}
        {% end %}
    }
}


}
