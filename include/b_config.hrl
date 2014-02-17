{# vim:se syntax=html.tornadotmpl: #}
{% autoescape None %}

-ifndef(LOGIC_CONFIG_H).
-define(LOGIC_CONFIG_H, true).


-record(hero_base, { {{", ".join(k for k in sorted(root.heroes[0])) }} }).
-record(item_base, { {{", ".join(k for k in sorted(root["items"][0])) }} }).


-endif.
