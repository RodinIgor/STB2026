using Xunit;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Reflection.Metadata;
using System.Reflection.PortableExecutable;

namespace STB2026.Tests
{
    public class HelloWorldCommandTests
    {
        private static readonly string DllPath;

        static HelloWorldCommandTests()
        {
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            DllPath = Path.GetFullPath(Path.Combine(
                baseDir, "..", "..", "..", "..",
                "STB2026", "bin", "Debug", "net8.0-windows", "STB2026.dll"));
        }

        private (List<string> types, List<string> interfaces, List<string> attributes, List<(string type, string method, int paramCount)> methods) ReadAssemblyMetadata()
        {
            var types = new List<string>();
            var interfaces = new List<string>();
            var attributes = new List<string>();
            var methods = new List<(string, string, int)>();

            using var stream = File.OpenRead(DllPath);
            using var peReader = new PEReader(stream);
            var mdReader = peReader.GetMetadataReader();

            foreach (var typeHandle in mdReader.TypeDefinitions)
            {
                var typeDef = mdReader.GetTypeDefinition(typeHandle);
                var ns = mdReader.GetString(typeDef.Namespace);
                var name = mdReader.GetString(typeDef.Name);
                var fullName = string.IsNullOrEmpty(ns) ? name : $"{ns}.{name}";
                types.Add(fullName);

                // Интерфейсы
                foreach (var ifaceHandle in typeDef.GetInterfaceImplementations())
                {
                    var iface = mdReader.GetInterfaceImplementation(ifaceHandle);
                    var ifaceName = GetFullName(mdReader, iface.Interface);
                    interfaces.Add($"{fullName}:{ifaceName}");
                }

                // Атрибуты
                foreach (var attrHandle in typeDef.GetCustomAttributes())
                {
                    var attr = mdReader.GetCustomAttribute(attrHandle);
                    var attrName = GetAttributeTypeName(mdReader, attr);
                    attributes.Add($"{fullName}:{attrName}");
                }

                // Методы
                foreach (var methodHandle in typeDef.GetMethods())
                {
                    var methodDef = mdReader.GetMethodDefinition(methodHandle);
                    var methodName = mdReader.GetString(methodDef.Name);
                    var sig = methodDef.DecodeSignature(new ParamCounter(), null);
                    methods.Add((fullName, methodName, sig.ParameterTypes.Length));
                }
            }

            return (types, interfaces, attributes, methods);
        }

        private static string GetFullName(MetadataReader reader, EntityHandle handle)
        {
            if (handle.Kind == HandleKind.TypeReference)
            {
                var typeRef = reader.GetTypeReference((TypeReferenceHandle)handle);
                var ns = reader.GetString(typeRef.Namespace);
                var name = reader.GetString(typeRef.Name);
                return string.IsNullOrEmpty(ns) ? name : $"{ns}.{name}";
            }
            if (handle.Kind == HandleKind.MemberReference)
            {
                var memberRef = reader.GetMemberReference((MemberReferenceHandle)handle);
                return GetFullName(reader, memberRef.Parent);
            }
            return handle.Kind.ToString();
        }

        private static string GetAttributeTypeName(MetadataReader reader, CustomAttribute attr)
        {
            if (attr.Constructor.Kind == HandleKind.MemberReference)
            {
                var memberRef = reader.GetMemberReference((MemberReferenceHandle)attr.Constructor);
                return GetFullName(reader, memberRef.Parent);
            }
            if (attr.Constructor.Kind == HandleKind.MethodDefinition)
            {
                var methodDef = reader.GetMethodDefinition((MethodDefinitionHandle)attr.Constructor);
                var typeDef = reader.GetTypeDefinition(methodDef.GetDeclaringType());
                var ns = reader.GetString(typeDef.Namespace);
                var name = reader.GetString(typeDef.Name);
                return string.IsNullOrEmpty(ns) ? name : $"{ns}.{name}";
            }
            return "";
        }

        // Простой декодер сигнатур — считаем только количество параметров
        private class ParamCounter : ISignatureTypeProvider<int, object?>
        {
            public int GetPrimitiveType(PrimitiveTypeCode typeCode) => 0;
            public int GetTypeFromDefinition(MetadataReader reader, TypeDefinitionHandle handle, byte rawTypeKind) => 0;
            public int GetTypeFromReference(MetadataReader reader, TypeReferenceHandle handle, byte rawTypeKind) => 0;
            public int GetTypeFromSpecification(MetadataReader reader, object? genericContext, TypeSpecificationHandle handle, byte rawTypeKind) => 0;
            public int GetSZArrayType(int elementType) => 0;
            public int GetArrayType(int elementType, ArrayShape shape) => 0;
            public int GetByReferenceType(int elementType) => 0;
            public int GetPointerType(int elementType) => 0;
            public int GetGenericInstantiation(int genericType, System.Collections.Immutable.ImmutableArray<int> typeArguments) => 0;
            public int GetGenericTypeParameter(object? genericContext, int index) => 0;
            public int GetGenericMethodParameter(object? genericContext, int index) => 0;
            public int GetModifiedType(int modifier, int unmodifiedType, bool isRequired) => 0;
            public int GetPinnedType(int elementType) => 0;
            public int GetFunctionPointerType(MethodSignature<int> signature) => 0;
        }

        // ===== ТЕСТЫ =====

        [Fact]
        public void AssemblyFile_Exists()
        {
            Assert.True(File.Exists(DllPath), $"STB2026.dll не найден: {DllPath}");
        }

        [Fact]
        public void Assembly_NameIsCorrect()
        {
            using var stream = File.OpenRead(DllPath);
            using var pe = new PEReader(stream);
            var reader = pe.GetMetadataReader();
            var asmDef = reader.GetAssemblyDefinition();
            Assert.Equal("STB2026", reader.GetString(asmDef.Name));
        }

        [Fact]
        public void Assembly_ContainsAppClass()
        {
            var (types, _, _, _) = ReadAssemblyMetadata();
            Assert.Contains("STB2026.App", types);
        }

        [Fact]
        public void Assembly_ContainsHelloWorldCommand()
        {
            var (types, _, _, _) = ReadAssemblyMetadata();
            Assert.Contains("STB2026.HelloWorldCommand", types);
        }

        [Fact]
        public void App_ImplementsIExternalApplication()
        {
            var (_, interfaces, _, _) = ReadAssemblyMetadata();
            Assert.Contains(interfaces,
                i => i == "STB2026.App:Autodesk.Revit.UI.IExternalApplication");
        }

        [Fact]
        public void App_HasOnStartupMethod()
        {
            var (_, _, _, methods) = ReadAssemblyMetadata();
            Assert.Contains(methods, m => m.type == "STB2026.App" && m.method == "OnStartup");
        }

        [Fact]
        public void App_HasOnShutdownMethod()
        {
            var (_, _, _, methods) = ReadAssemblyMetadata();
            Assert.Contains(methods, m => m.type == "STB2026.App" && m.method == "OnShutdown");
        }

        [Fact]
        public void HelloWorldCommand_ImplementsIExternalCommand()
        {
            var (_, interfaces, _, _) = ReadAssemblyMetadata();
            Assert.Contains(interfaces,
                i => i == "STB2026.HelloWorldCommand:Autodesk.Revit.UI.IExternalCommand");
        }

        [Fact]
        public void HelloWorldCommand_HasTransactionAttribute()
        {
            var (_, _, attributes, _) = ReadAssemblyMetadata();
            Assert.Contains(attributes,
                a => a.StartsWith("STB2026.HelloWorldCommand:") &&
                     a.Contains("TransactionAttribute"));
        }

        [Fact]
        public void HelloWorldCommand_HasExecuteMethod()
        {
            var (_, _, _, methods) = ReadAssemblyMetadata();
            var execute = methods.FirstOrDefault(
                m => m.type == "STB2026.HelloWorldCommand" && m.method == "Execute");
            Assert.NotEqual(default, execute);
            Assert.Equal(3, execute.paramCount);
        }

        [Fact]
        public void App_HasCorrectNamespace()
        {
            var (types, _, _, _) = ReadAssemblyMetadata();
            Assert.Contains("STB2026.App", types);
        }

        [Fact]
        public void HelloWorldCommand_HasCorrectNamespace()
        {
            var (types, _, _, _) = ReadAssemblyMetadata();
            Assert.Contains("STB2026.HelloWorldCommand", types);
        }
    }
}