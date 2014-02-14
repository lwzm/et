{# vim:se syntax=html.tornadotmpl: #}
{% autoescape None %}

-module(b_config).
-export([get/1, set/2]).

set(K, V) -> {todo, K, V}.

{% for x in root.hero_strengthen %}
get({ hero_strengthen, {{ x.main_star }}, {{ x.sub_star }} }) -> { {{ x.rate }}, {{ x.point }} };
{% end %}

{% set key_head = "hero_strengthen_gold_cost_" %}
{% for k in sorted(root) %}
{% if k.startswith(key_head) %}
{% set i = int(k[len(key_head):]) %}
{% for x in root[k] %}
get({ hero_strengthen_gold_cost, {{ i }}, {{ x.lv }} }) -> {{ x.gold }};
{% end %}
{% end %}
{% end %}

{% set ignore = {"id", "name", "icon", "asset", "grow"} %}
{% for x in root.heroes %}
get({ heroes, {{ x.id }}, {{ x.lv }} }) -> { hero_base, {{", ". join(repr(v) for k, v in sorted(x.items()) if k not in ignore)}} };
{% end %}

get(_) -> "Not Found".
