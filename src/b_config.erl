{# vim:se syntax=html.tornadotmpl: #}
{% autoescape None %}
-module(b_config).
-compile(export_all).

{% for x in root.hero_strengthen %}
hero_strengthen({{ x.main_star }}, {{ x.sub_star }}) -> { {{ x.rate }}, {{ x.point }} };
{% end %}
hero_strengthen('end', 'end') -> 'end'.

{% set key_head = "hero_strengthen_gold_cost_" %}
{% for k in sorted(root) %}
{% if k.startswith(key_head) %}
{% set i = int(k[len(key_head):]) %}
{% for x in root[k] %}
hero_strengthen_gold_cost({{ i }}, {{ x.level }}) -> {{ x.gold_cost }};
{% end %}
{% end %}
{% end %}
hero_strengthen_gold_cost('end', 'end') -> 'end'.

{% set ignore = {"id", "name", "icon", "asset"} %}
{% for x in root.heroes %}
heroes({{ x.id }}) -> { hero_base, {{", ". join(repr(v) for k, v in sorted(x.items()) if k not in ignore)}} };
{% end %}
heroes('end') -> 'end'.
