using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace RibbonButtons
{
    public static class AssemblyLoader
    {
        public class MethodIdentifier
        {
            public string AssemblyPath { get; set; }
            public string ClassName { get; set; }
            public string MethodName { get; set; }  // public static method

            public override int GetHashCode()
            {
                return AssemblyPath.GetHashCode() ^ ClassName.GetHashCode() ^ MethodName.GetHashCode();
            }

            public override bool Equals(object obj)
            {
                MethodIdentifier anotherId = (MethodIdentifier)obj;
                return this.AssemblyPath == anotherId.AssemblyPath
                    && this.ClassName == anotherId.ClassName
                    && this.MethodName == anotherId.MethodName;
            }
        }

        private static object _locker = new object();

        private static Dictionary<string, Assembly> Assemblies = new Dictionary<string, Assembly>();
        private static Dictionary<MethodIdentifier, MethodInfo> Methods  = new Dictionary<MethodIdentifier,MethodInfo>();

        public static Assembly LoadAssembly(string assemblyPath)
        {
            if (!Assemblies.ContainsKey(assemblyPath))
            {
                lock (_locker)
                {
                    if (!Assemblies.ContainsKey(assemblyPath))
                    {
                        Assembly assembly = Assembly.LoadFile(assemblyPath);

                        Assemblies.Add(assemblyPath, assembly);
                    }
                }
            }

            return Assemblies[assemblyPath];
        }

        public static void InvokeMethod(MethodIdentifier methodId, string args)
        {
            if (!Methods.ContainsKey(methodId))
            {
                lock (_locker)
                {
                    if (!Methods.ContainsKey(methodId))
                    {
                        Assembly assembly = LoadAssembly(methodId.AssemblyPath);
                        Type programType = assembly.GetType(methodId.ClassName);
                        MethodInfo method = programType.GetMethod(methodId.MethodName, BindingFlags.Static | BindingFlags.Public);
                        
                        Methods.Add(methodId, method);
                    }
                }
            }

            Methods[methodId].Invoke(null, new object[] { new string[] { args } });
        }
    }
}
