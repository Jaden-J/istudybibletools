using System;
using System.Xml;
using System.Runtime.Serialization.Formatters.Binary;
using System.IO;
using System.Text;
using System.Runtime.Serialization;
using System.Reflection;

namespace BibleCommon.Helpers
{    
    /// <summary>
    /// Serialize/Deserialize any object from/to any file or serialize/Deserialize it
    /// into/From bytes array
    /// </summary>
    public static class BinarySerializerHelper
    {
        #region Constants       
        /// <summary>
        /// Category for Exceptions thrown
        /// </summary>
        private const string CATEGORY = "BinarySerializerHelper::{0}";

        #endregion

        #region Members
        /// <summary>
        /// Singleton object of BinaryFormatter
        /// </summary>
        private static BinaryFormatter _formatter;

        public static BinaryFormatter Formatter
        {
            get
            {
                if (_formatter == null)
                {
                    _formatter = new BinaryFormatter();
                    _formatter.Binder = new AllowAllAssemblyVersionsDeserializationBinder();
                }

                return _formatter;
            }
        }
        
        #endregion

        #region Methods
        /// <summary>
        /// Serialize Any object to a file
        /// </summary>
        /// <param name="currentObject"><see cref="System.Object"/></param>
        /// <param name="filename">file name and where this file will be saved</param>
        /// <remarks>The application should have enough permissions to
        /// access/create a file on the targeted machine</remarks>
        public static void Serialize(object currentObject, string filename)
        {
            if(null == currentObject)
                throw new ArgumentNullException(string.Format(
                    CATEGORY, "Serialize, Passed Object to Serialize Is Null"));

            if(string.IsNullOrEmpty(filename))
                throw new ArgumentNullException(string.Format(
                    CATEGORY, "Serialize,Cannot Serialize object, Filename Is Null"));

            using(Stream fileStream = new FileStream(filename, FileMode.Create))
                Formatter.Serialize(fileStream, currentObject);

        }
        /// <summary>
        /// Serialize any object to a bytes array
        /// </summary>
        /// <param name="currentObject"><see cref="System.Object"/></param>
        /// <returns><see cref="System.Byte"/></returns>
        public static byte[] Serialize(object currentObject)
        {
            if(null == currentObject)
                throw new ArgumentNullException(string.Format(
                    CATEGORY, "Serialize, Object to Serialize Is Null"));

            byte[] binaryData = null;
            
            using(Stream memoryStream = new MemoryStream())            
            {
                Formatter.Serialize(memoryStream, currentObject);        
                memoryStream.Position = 0;
                binaryData = new byte[memoryStream.Length];
                memoryStream.Read(binaryData, 0, Convert.ToInt32(memoryStream.Length));
            }

            return binaryData;
        }
        /// <summary>
        /// Deserialize any binary data to an object from a bytes array 
        /// </summary>        
        /// <returns><see cref="System.byte"/></returns>
        /// <remarks>Client should be aware of the object type</remarks>
        public static object Deserialize(byte[] binaryData)
        {
            if(null == binaryData || binaryData.Length == 0)
                throw new ArgumentNullException(string.Format(
                    CATEGORY, "Deserialize, binaryData Is Null or Empty"));

            object deserializedObject = null;

            using(Stream memoryStream = new MemoryStream())
            {            
                memoryStream.Write(binaryData, 0, binaryData.Length);
                memoryStream.Position = 0;
                deserializedObject = Formatter.Deserialize(memoryStream);
            }

            return deserializedObject;
        }
        /// <summary>
        /// Deserialize any binary data to an object from a file 
        /// </summary>
        /// <param name="filename">where the binary data is located</param>
        /// <returns><see cref="System.Object"/></returns>
        /// <remarks>Client should be aware of the object type</remarks>
        public static object Deserialize(string filename)
        {
            if(string.IsNullOrEmpty(filename))
                throw new ArgumentNullException(string.Format(
                    CATEGORY, "Deserialize, Filename Is Null"));

            object deserializedObject = null;

            using (MemoryStream ms = new MemoryStream(File.ReadAllBytes(filename)))                        
                deserializedObject = Formatter.Deserialize(ms);            
            

            return deserializedObject;
        } 
        #endregion
    }

    sealed class AllowAllAssemblyVersionsDeserializationBinder : System.Runtime.Serialization.SerializationBinder
    {
        public override Type BindToType(string assemblyName, string typeName)
        {
            Type typeToDeserialize = null;

            String currentAssembly = Assembly.GetExecutingAssembly().FullName;

            // In this case we are always using the current assembly
            assemblyName = currentAssembly;

            // Get the type using the typeName and assemblyName
            typeToDeserialize = Type.GetType(String.Format("{0}, {1}",
                typeName, assemblyName));

            return typeToDeserialize;
        }
    }
}
