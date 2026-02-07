using Xunit;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace STB2026.Tests
{
    /// <summary>
    /// Структурные тесты HelloWorldCommand и App.
    /// Используют MetadataLoadContext для проверки без загрузки Revit API.
    /// </summary>
    public class HelloWorldCommandTests : IDisposable
    {
        private readonly MetadataLoadContext _mlc;
        private readonly Assembly _asm;

        public HelloWorldCommandTests()
        {
            string testDir = Path.GetDirectoryName(typeof(HelloWorldCommandTests).Assembly.Location)!;
            string mainDll = Path.Combine(testDir, "STB2026.dll");

            if (!File.Exists(mainDll))
            {
                string projDir = Path.GetFullPath(Path.Combine(testDir, "..", "..", "..", "..", "STB2026", "bin", "Debug", "net8.0-windows"));
                mainDll = Path.Combine(projDir, "STB2026.dll");
            }

            if (!File.Exists(mainDll))
                throw new FileNotFoundException("STB2026.dll не найдена. Соберите основной проект.", mainDll);

            string revitDir = @"C:\Program Files\Autodesk\Revit 2025";
            string rtDir = RuntimeEnvironment.GetRuntimeDirectory();

            var allPaths = new[] {
                mainDll,
                Path.Combine(revitDir, "RevitAPI.dll"),
                Path.Combine(revitDir, "RevitAPIUI.dll"),
                typeof(object).Assembly.Location,
                Path.Combine(rtDir, "System.Runtime.dll"),
                Path.Combine(rtDir, "System.Collections.dll"),
                Path.Combine(rtDir, "netstandard.dll"),
            }.Where(File.Exists);

            _mlc = new MetadataLoadContext(new PathAssemblyResolver(allPaths));
            _asm = _mlc.LoadFromAssemblyPath(mainDll);
        }

        public void Dispose() => _mlc?.Dispose();

        // ========================================
        // HelloWorldCommand
        // ========================================

        [Fact]
        public void HelloWorldCommand_ImplementsIExternalCommand()
        {
            var type = _asm.GetType("STB2026.HelloWorldCommand");
            Assert.NotNull(type);
            Assert.True(
                type!.GetInterfaces().Any(i => i.FullName == "Autodesk.Revit.UI.IExternalCommand"),
                "HelloWorldCommand должен реализовывать IExternalCommand");
        }

        [Fact]
        public void HelloWorldCommand_HasTransactionAttribute()
        {
            var type = _asm.GetType("STB2026.HelloWorldCommand")!;
            var attr = type.CustomAttributes
                .FirstOrDefault(a => a.AttributeType.Name == "TransactionAttribute");
            Assert.NotNull(attr);
        }

        [Fact]
        public void HelloWorldCommand_TransactionModeIsManual()
        {
            var type = _asm.GetType("STB2026.HelloWorldCommand")!;
            var attr = type.CustomAttributes
                .First(a => a.AttributeType.Name == "TransactionAttribute");
            // TransactionMode enum: проверяем имя значения
            var modeValue = (int)attr.ConstructorArguments[0].Value!;
            var modeType = attr.ConstructorArguments[0].ArgumentType;
            var modeName = Enum.GetName(modeType, modeValue);
            Assert.Equal("Manual", modeName);
        }

        [Fact]
        public void HelloWorldCommand_HasExecuteMethod()
        {
            var type = _asm.GetType("STB2026.HelloWorldCommand")!;
            var method = type.GetMethod("Execute");
            Assert.NotNull(method);
            Assert.Equal("Result", method!.ReturnType.Name);
            Assert.Equal(3, method.GetParameters().Length);
        }

        [Fact]
        public void HelloWorldCommand_HasDefaultConstructor()
        {
            var type = _asm.GetType("STB2026.HelloWorldCommand")!;
            Assert.True(type.GetConstructors().Any(c => c.GetParameters().Length == 0));
        }

        // ========================================
        // App
        // ========================================

        [Fact]
        public void App_ImplementsIExternalApplication()
        {
            var type = _asm.GetType("STB2026.App");
            Assert.NotNull(type);
            Assert.True(
                type!.GetInterfaces().Any(i => i.FullName == "Autodesk.Revit.UI.IExternalApplication"),
                "App должен реализовывать IExternalApplication");
        }

        [Fact]
        public void App_HasOnStartupMethod()
        {
            var type = _asm.GetType("STB2026.App")!;
            var method = type.GetMethod("OnStartup");
            Assert.NotNull(method);
            Assert.Equal("Result", method!.ReturnType.Name);
        }

        [Fact]
        public void App_HasOnShutdownMethod()
        {
            var type = _asm.GetType("STB2026.App")!;
            var method = type.GetMethod("OnShutdown");
            Assert.NotNull(method);
            Assert.Equal("Result", method!.ReturnType.Name);
        }

        [Fact]
        public void App_HasDefaultConstructor()
        {
            var type = _asm.GetType("STB2026.App")!;
            Assert.True(type.GetConstructors().Any(c => c.GetParameters().Length == 0));
        }

        // ========================================
        // Сборка
        // ========================================

        [Fact]
        public void Assembly_ContainsExpectedTypes()
        {
            Assert.NotNull(_asm.GetType("STB2026.App"));
            Assert.NotNull(_asm.GetType("STB2026.HelloWorldCommand"));
        }

        [Fact]
        public void Assembly_NameIsCorrect()
        {
            Assert.Equal("STB2026", _asm.GetName().Name);
        }

        [Fact]
        public void Assembly_HasCorrectNamespace()
        {
            Assert.Equal("STB2026", _asm.GetType("STB2026.App")!.Namespace);
            Assert.Equal("STB2026", _asm.GetType("STB2026.HelloWorldCommand")!.Namespace);
        }
    }
}
