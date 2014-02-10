{# vim:se syntax=html.tornadotmpl: #}
{% autoescape None %}

-module(b_proto).
-export([en/1, de/1]).
-include_lib("eunit/include/eunit.hrl").

{% set protobuffs = root.protobuffs %}

{% for hrl in set(i.module for i in protobuffs) %}
-include("{{ hrl }}_pb.hrl").
{% end %}

-define(INT, 16/unsigned-big-integer).

{% for i in protobuffs %}
en(#{{ i.message }}{}=M) -> <<{{ i.id }}:?INT, (erlang:iolist_to_binary({{ i.module }}_pb:encode_{{ i.message }}(M)))/binary>>;
{% end %}
en(_) -> error.

{% for i in protobuffs %}
de(<<{{ i.id }}:?INT, B/binary>>) -> {{ i.module }}_pb:decode_{{ i.message }}(B);
{% end %}
de(_) -> error.
