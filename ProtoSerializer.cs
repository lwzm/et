using System;
using System.Collections.Generic;
using System.IO;
using ProtoBuf;


{% set protobuffs = root["protobuffs"] %}

public enum ProtoNameIds {
{% for i in protobuffs %}
    {{ i["message"].upper() }} = {{ i["id"] }},
{% end %}
}

public class ProtoIdNames
{
    public static Dictionary<ProtoNameIds, string> protoIdNames = new Dictionary<ProtoNameIds, string>();
    static ProtoIdNames() {
    {% for i in protobuffs %}
        protoIdNames.Add(ProtoNameIds.{{ i["message"].upper() }}, "network.{{ i["message"] }}");
    {% end %}
    }
}

public class ProtoSerializer
{
    public static Object ParseFrom(ProtoNameIds protoType, Stream stream)
    {
        switch (protoType) {
        {% for i in protobuffs %}
            case ProtoNameIds.{{ i["message"].upper() }}: return Serializer.Deserialize<network.{{ i["message"] }}>(stream);
        {% end %}
            default: break;
        }
        return null;
    }

    public static void Serialize(ProtoNameIds protoType, Stream stream, Object proto)
    {
        switch (protoType)
        {
        {% for i in protobuffs %}
            case ProtoNameIds.{{ i["message"].upper() }}: Serializer.Serialize(stream, (network.{{ i["message"] }})proto); break;
        {% end %}
            default: break;
        }
    }
}
