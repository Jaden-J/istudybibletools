using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Polenter.Serialization;
using Polenter.Serialization.Core;
using System.IO;
using System.Reflection;

namespace BibleCommon.Helpers
{
    public static class SharpSerializationHelper
    {
        public static T Deserialize<T>(string filePath)
        {
            var serializer = new SharpSerializer(true);
            return (T)serializer.Deserialize(filePath);
        }

        public static void Serialize(object data, string filePath)
        {
            var serializer = new SharpSerializer(true);
            serializer.Serialize(data, filePath);
        }
    }
}
