// --------------------------------------------------------------------------------------------------------------------
// <copyright file="AsposeCellsLicenseTest.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.AsposeCells.Test
{
    using System;

    using FakeItEasy;

    using FluentAssertions;

    using Xunit;

    public static class AsposeCellsLicenseTest
    {
        [Fact]
        public static void Constructor___Should_throw_ArgumentNullException___When_parameter_licenseXml_is_null()
        {
            // Arrange, Act
            var actual = Record.Exception(() => new AsposeCellsLicense(null));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("licenseXml");
        }

        [Fact]
        public static void Constructor___Should_throw_ArgumentException___When_parameter_licenseXml_is_white_space()
        {
            // Arrange, Act
            var actual = Record.Exception(() => new AsposeCellsLicense(" \r\n "));

            // Assert
            actual.Should().BeOfType<ArgumentException>();
            actual.Message.Should().Contain("licenseXml");
            actual.Message.Should().Contain("white space");
        }

        [Fact]
        public static void Register___Should_throw_InvalidOperationException___When_licenseXml_is_invalid_or_corrupt()
        {
            // Arrange
            var systemUnderTest1 = new AsposeCellsLicense(A.Dummy<string>());
            var systemUnderTest2 = new AsposeCellsLicense("<License></License>");

            // Act
            var actual1 = Record.Exception(() => systemUnderTest1.Register());
            var actual2 = Record.Exception(() => systemUnderTest2.Register());

            // Assert
            actual1.Should().BeOfType<InvalidOperationException>();
            actual1.Message.Should().Contain("LicenseXml is invalid or corrupt");

            actual2.Should().BeOfType<InvalidOperationException>();
            actual2.Message.Should().Contain("LicenseXml is invalid or corrupt");
        }
    }
}
