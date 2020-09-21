﻿// --------------------------------------------------------------------------------------------------------------------
// <auto-generated>
//   Generated using OBeautifulCode.CodeGen.ModelObject (1.0.106.0)
// </auto-generated>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.Test
{
    using global::System.CodeDom.Compiler;
    using global::System.Diagnostics.CodeAnalysis;

    using global::FakeItEasy;
    
    using global::OBeautifulCode.Assertion.Recipes;

    using global::Xunit;

    [ExcludeFromCodeCoverage]
    [GeneratedCode("OBeautifulCode.CodeGen.ModelObject", "1.0.106.0")]
    public static partial class ExcelDummyFactoryTest
    {
        [Fact]
        public static void ExcelDummyFactory___Should_derive_from_DefaultExcelDummyFactory___When_reflecting()
        {
            // Arrange
            var dummyFactoryType = typeof(ExcelDummyFactory);

            var defaultDummyFactoryType = typeof(DefaultExcelDummyFactory);

            // Act, Assert
            defaultDummyFactoryType.GetInterface(nameof(IDummyFactory)).AsTest().Must().NotBeNull();

            dummyFactoryType.BaseType.AsTest().Must().BeEqualTo(defaultDummyFactoryType);
        }
    }
}