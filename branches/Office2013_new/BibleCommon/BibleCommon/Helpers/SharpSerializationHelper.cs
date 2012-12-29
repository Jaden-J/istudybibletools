﻿using System;
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
            using (var ms = new MemoryStream(File.ReadAllBytes(filePath)))
            {
                return (T)serializer.Deserialize(ms);
            }
        }

        public static void Serialize(object data, string filePath)
        {
            var serializer = new SharpSerializer(true);
            serializer.Serialize(data, filePath);
        }
    }
}