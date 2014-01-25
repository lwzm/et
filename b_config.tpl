{# vim:se syntax=html.tornadotmpl: #}
{% autoescape None %}
-module(b_config).
-compile(export_all).

{% for x in root.hero_strengthen %}
hero_strengthen({{ x.main_star }}, {{ x.sub_star }}) -> { {{ x.rate }}, {{ x.point }} };
{% end %}
hero_strengthen('end', 'end') -> 'end'.

{% for k in sorted(root) %}
{% if k.startswith("hero_strengthen_gold_cost_") %}
{% set i = int(k[len("hero_strengthen_gold_cost_"):]) %}
{% for x in root[k] %}
hero_strengthen_gold_cost({{ i }}, {{ x.level }}) -> {{ x.gold_cost }};
{% end %}
{% end %}
{% end %}
hero_strengthen_gold_cost('end', 'end') -> 'end'.
