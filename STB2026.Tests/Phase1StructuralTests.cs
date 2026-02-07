using Xunit;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace STB2026.Tests
{
    /// <summary>
    /// Структурные тесты Фазы 1: команды, сервисы, модели, нормы СП 60.
    /// Все тесты через MetadataLoadContext — без загрузки Revit API.
    /// </summary>
    public class Phase1StructuralTests : IDisposable
    {
        private readonly MetadataLoadContext _mlc;
        private readonly Assembly _asm;

        public Phase1StructuralTests()
        {
            string testDir = Path.GetDirectoryName(typeof(Phase1StructuralTests).Assembly.Location)!;
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
                Path.Combine(rtDir, "System.Linq.dll"),
                Path.Combine(rtDir, "System.ObjectModel.dll"),
                Path.Combine(rtDir, "netstandard.dll"),
            }.Where(File.Exists);

            _mlc = new MetadataLoadContext(new PathAssemblyResolver(allPaths));
            _asm = _mlc.LoadFromAssemblyPath(mainDll);
        }

        public void Dispose() => _mlc?.Dispose();

        // ═══════════════════════════════════════
        // Вспомогательные
        // ═══════════════════════════════════════

        private Type MustGet(string fullName)
        {
            var t = _asm.GetType(fullName);
            Assert.True(t != null, $"Тип {fullName} не найден в сборке");
            return t!;
        }

        // ═══════════════════════════════════════
        // Команды
        // ═══════════════════════════════════════

        [Theory]
        [InlineData("STB2026.Commands.TagDuctsCommand")]
        [InlineData("STB2026.Commands.VelocityCheckCommand")]
        [InlineData("STB2026.Commands.WallIntersectionsCommand")]
        [InlineData("STB2026.Commands.SystemValidationCommand")]
        public void Command_ImplementsIExternalCommand(string typeName)
        {
            var cmd = MustGet(typeName);
            Assert.True(
                cmd.GetInterfaces().Any(i => i.FullName == "Autodesk.Revit.UI.IExternalCommand"),
                $"{typeName} должен реализовывать IExternalCommand");
        }

        [Theory]
        [InlineData("STB2026.Commands.TagDuctsCommand")]
        [InlineData("STB2026.Commands.VelocityCheckCommand")]
        [InlineData("STB2026.Commands.WallIntersectionsCommand")]
        [InlineData("STB2026.Commands.SystemValidationCommand")]
        public void Command_HasTransactionAttribute(string typeName)
        {
            var cmd = MustGet(typeName);
            var attr = cmd.CustomAttributes.FirstOrDefault(a => a.AttributeType.Name == "TransactionAttribute");
            Assert.NotNull(attr);
            // Проверяем имя значения enum (Manual), а не числовое значение
            var modeValue = (int)attr!.ConstructorArguments[0].Value!;
            var modeType = attr.ConstructorArguments[0].ArgumentType;
            var modeName = Enum.GetName(modeType, modeValue);
            Assert.Equal("Manual", modeName);
        }

        [Theory]
        [InlineData("STB2026.Commands.TagDuctsCommand")]
        [InlineData("STB2026.Commands.VelocityCheckCommand")]
        [InlineData("STB2026.Commands.WallIntersectionsCommand")]
        [InlineData("STB2026.Commands.SystemValidationCommand")]
        public void Command_HasExecuteMethod(string typeName)
        {
            var cmd = MustGet(typeName);
            var method = cmd.GetMethod("Execute");
            Assert.NotNull(method);
            Assert.Equal("Result", method!.ReturnType.Name);
            Assert.Equal(3, method.GetParameters().Length);
        }

        [Theory]
        [InlineData("STB2026.Commands.TagDuctsCommand")]
        [InlineData("STB2026.Commands.VelocityCheckCommand")]
        [InlineData("STB2026.Commands.WallIntersectionsCommand")]
        [InlineData("STB2026.Commands.SystemValidationCommand")]
        public void Command_HasDefaultConstructor(string typeName)
        {
            var cmd = MustGet(typeName);
            Assert.True(cmd.GetConstructors().Any(c => c.GetParameters().Length == 0));
        }

        // ═══════════════════════════════════════
        // VelocityNorms — структура
        // ═══════════════════════════════════════

        [Fact]
        public void VelocityNorms_HasCheckVelocitySimpleMethod()
        {
            var vn = MustGet("STB2026.Models.VelocityNorms");
            var m = vn.GetMethod("CheckVelocitySimple", BindingFlags.Public | BindingFlags.Static);
            Assert.NotNull(m);
            Assert.Equal(3, m!.GetParameters().Length);
        }

        [Fact]
        public void VelocityNorms_HasGetRangeMethod()
        {
            var vn = MustGet("STB2026.Models.VelocityNorms");
            var m = vn.GetMethod("GetRange", BindingFlags.Public | BindingFlags.Static);
            Assert.NotNull(m);
            Assert.Equal(2, m!.GetParameters().Length);
        }

        [Fact]
        public void VelocityNorms_HasVelocityStatusEnum()
        {
            var status = _asm.GetType("STB2026.Models.VelocityNorms+VelocityStatus");
            Assert.NotNull(status);
            Assert.True(status!.IsEnum);
            var names = Enum.GetNames(status);
            Assert.Contains("Normal", names);
            Assert.Contains("Warning", names);
            Assert.Contains("Exceeded", names);
            Assert.Contains("NoData", names);
        }

        [Fact]
        public void VelocityNorms_HasSpeedRangeClass()
        {
            var sr = _asm.GetType("STB2026.Models.VelocityNorms+SpeedRange");
            Assert.NotNull(sr);
            Assert.NotNull(sr!.GetProperty("Min"));
            Assert.NotNull(sr.GetProperty("Max"));
            Assert.NotNull(sr.GetProperty("Description"));
        }

        [Fact]
        public void VelocityNorms_HasStaticTables()
        {
            var vn = MustGet("STB2026.Models.VelocityNorms");
            // SupplyPublic и ExhaustPublic — это static readonly поля, а не свойства
            Assert.NotNull(vn.GetField("SupplyPublic", BindingFlags.Public | BindingFlags.Static));
            Assert.NotNull(vn.GetField("ExhaustPublic", BindingFlags.Public | BindingFlags.Static));
            Assert.NotNull(vn.GetField("SupplyIndustrial", BindingFlags.Public | BindingFlags.Static));
            Assert.NotNull(vn.GetField("ExhaustIndustrial", BindingFlags.Public | BindingFlags.Static));
        }

        [Fact]
        public void VelocityNorms_HasWarningThreshold()
        {
            var vn = MustGet("STB2026.Models.VelocityNorms");
            Assert.NotNull(vn.GetField("WarningThreshold", BindingFlags.Public | BindingFlags.Static));
        }

        [Fact]
        public void VelocityNorms_HasDefaultMaxConstants()
        {
            var vn = MustGet("STB2026.Models.VelocityNorms");
            Assert.NotNull(vn.GetField("DefaultMaxSupply", BindingFlags.Public | BindingFlags.Static));
            Assert.NotNull(vn.GetField("DefaultMaxExhaust", BindingFlags.Public | BindingFlags.Static));
        }

        // ═══════════════════════════════════════
        // Модели результатов
        // ═══════════════════════════════════════

        [Fact]
        public void TaggingResult_HasExpectedProperties()
        {
            var tr = MustGet("STB2026.Services.TaggingResult");
            Assert.NotNull(tr.GetProperty("Tagged"));
            Assert.NotNull(tr.GetProperty("AlreadyTagged"));
            Assert.NotNull(tr.GetProperty("Skipped"));
            Assert.NotNull(tr.GetProperty("Errors"));
        }

        [Fact]
        public void VelocityCheckResult_HasExpectedProperties()
        {
            var vcr = MustGet("STB2026.Services.VelocityCheckResult");
            Assert.NotNull(vcr.GetProperty("Total"));
            Assert.NotNull(vcr.GetProperty("Normal"));
            Assert.NotNull(vcr.GetProperty("Warning"));
            Assert.NotNull(vcr.GetProperty("Exceeded"));
            Assert.NotNull(vcr.GetProperty("NoData"));
        }

        [Fact]
        public void ValidationResult_HasExpectedProperties()
        {
            var vr = MustGet("STB2026.Services.ValidationResult");
            Assert.NotNull(vr.GetProperty("ZeroFlowIds"));
            Assert.NotNull(vr.GetProperty("DisconnectedIds"));
            Assert.NotNull(vr.GetProperty("NoSystemIds"));
            Assert.NotNull(vr.GetProperty("HasErrors"));
        }

        [Fact]
        public void IntersectionPoint_HasExpectedProperties()
        {
            var ip = MustGet("STB2026.Services.IntersectionPoint");
            foreach (var prop in new[] { "DuctId", "WallId", "X", "Y", "Z", "SystemName", "DuctSize" })
                Assert.NotNull(ip.GetProperty(prop));
        }

        [Fact]
        public void WallIntersectionResult_HasExpectedProperties()
        {
            var wir = MustGet("STB2026.Services.WallIntersectionResult");
            Assert.NotNull(wir.GetProperty("Intersections"));
            Assert.NotNull(wir.GetProperty("UniqueDucts"));
            Assert.NotNull(wir.GetProperty("UniqueWalls"));
        }

        // ═══════════════════════════════════════
        // Сервисы — существование
        // ═══════════════════════════════════════

        [Theory]
        [InlineData("STB2026.Services.DuctTaggerService")]
        [InlineData("STB2026.Services.VelocityCheckerService")]
        [InlineData("STB2026.Services.WallIntersectorService")]
        [InlineData("STB2026.Services.SystemValidatorService")]
        public void Service_Exists(string typeName)
        {
            MustGet(typeName);
        }

        // ═══════════════════════════════════════
        // Полнота сборки
        // ═══════════════════════════════════════

        [Fact]
        public void Assembly_ContainsAllPhase1Types()
        {
            var expected = new[] {
                "STB2026.App", "STB2026.HelloWorldCommand",
                "STB2026.Commands.TagDuctsCommand", "STB2026.Commands.VelocityCheckCommand",
                "STB2026.Commands.WallIntersectionsCommand", "STB2026.Commands.SystemValidationCommand",
                "STB2026.Services.DuctTaggerService", "STB2026.Services.VelocityCheckerService",
                "STB2026.Services.WallIntersectorService", "STB2026.Services.SystemValidatorService",
                "STB2026.Models.VelocityNorms",
            };
            foreach (var t in expected)
                Assert.True(_asm.GetType(t) != null, $"Тип {t} отсутствует");
        }

        [Fact]
        public void Assembly_NameIsCorrect()
        {
            Assert.Equal("STB2026", _asm.GetName().Name);
        }
    }
}
