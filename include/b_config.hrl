{# vim:se syntax=html.tornadotmpl: #}
{% autoescape None %}

{% set ignore = {"id", "name", "icon", "asset", "grow"} %}
-record(hero_base, { {{", ".join(k for k in sorted(root.heroes[0]) if k not in ignore) }} }).
