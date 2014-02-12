using System;
using System.Collections.Generic;
using System.IO;
using ProtoBuf;


{% set protobuffs = root["protobuffs"] %}

public class ProtoIdNames
{
    public static Dictionary<int, ProtoIdName> protoIdNames = new Dictionary<int, ProtoIdName>();
    static ProtoIdNames() {
    {% for i in protobuffs %}
        protoIdNames.Add({{ i["id"] }}, "network.{{ i["message"] }}"));
    {% end %}
    }
}

public class ProtoSerializer
{
    public static Object ParseFrom(int protoType, Stream stream)
    {
        switch (protoType) {
        {% for i in protobuffs %}
            case {{ i["id"] }}: return Serializer.Deserialize<network.{{ i["message"] }}>(stream);
        {% end %}
            default: break;
        }
        return null;
    }

    public static void Serialize(int protoType, Stream stream, Object proto)
    {
        switch (protoType)
        {
        {% for i in protobuffs %}
            case {{ i["id"] }}: Serializer.Serialize(stream, (network.{{ i["message"] }})proto); break;
        {% end %}
            default: break;
        }
    }
}
