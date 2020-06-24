// --------------------------------------------------------------------------------------------------------------------
// <copyright file="WorksheetProtectionTest.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.Test
{
    using FakeItEasy;

    using OBeautifulCode.Assertion.Recipes;

    using Xunit;

    public static partial class WorksheetProtectionTest
    {
        [Fact]
        public static void ToWorkbookProtection___Should_return_corresponding_WorkbookProtection___When_called()
        {
            // Arrange
            var systemUnderTest = A.Dummy<WorksheetProtection>();

            var expected = new WorkbookProtection
            {
                ClearTextPassword = systemUnderTest.ClearTextPassword,
            };

            // Act
            var actual = systemUnderTest.ToWorkbookProtection();

            // Assert
            actual.AsTest().Must().BeEqualTo(expected);
        }
    }
}
