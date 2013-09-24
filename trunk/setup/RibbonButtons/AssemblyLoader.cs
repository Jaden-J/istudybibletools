using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;

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
                        LoadAssemblyInternal(assemblyPath);                        

                        //LoadSatelliteAssemblies(assemblyPath);
                    }
                }
            }

            return Assemblies[assemblyPath];
        }

        private static void LoadAssemblyInternal(string assemblyPath)
        {
            try
            {
                var assembly = assemblyPath.EndsWith(".exe") ? Assembly.LoadFile(assemblyPath) : Assembly.LoadFrom(assemblyPath);

                Assemblies.Add(assemblyPath, assembly);
            }
            catch (FileNotFoundException ex)
            {
                MessageBox.Show(ex.Message + ": " + assemblyPath);
            }
        }

        //private static void LoadSatelliteAssemblies(string assemblyPath)
        //{

        //    foreach (var sap in Directory.GetFiles(Directory.GetParent(assemblyPath).ToString(),
        //                                    string.Format("{0}.resources.dll", Path.GetFileNameWithoutExtension(assemblyPath)), SearchOption.AllDirectories))
        //    {
        //        LoadAssemblyInternal(sap);
        //    }
        //}

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
