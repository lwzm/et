{# vim:se syntax=html.tornadotmpl: #}
{% autoescape None %}

-module(b_proto).
-export([en/1, de/1]).
-include_lib("eunit/include/eunit.hrl").

{% set protobuffs = root.protobuffs %}

{% for hrl in sorted(set(i.module for i in protobuffs)) %}
-include_lib("network_proto/include/{{ hrl }}_pb.hrl").
{% end %}

-define(INT, 16/unsigned-big-integer).

{% for i in protobuffs %}
en(#{{ i.message.lower() }}{}=M) -> {ok, <<{{ i.id }}:?INT, (erlang:iolist_to_binary({{ i.module }}_pb:encode_{{ i.message.lower() }}(M)))/binary>>};
{% end %}
en(_) -> error.

{% for i in protobuffs %}
de(<<{{ i.id }}:?INT, B/binary>>) -> {ok, {{ i.module }}_pb:decode_{{ i.message.lower() }}(B)};
{% end %}
de(_) -> error.
