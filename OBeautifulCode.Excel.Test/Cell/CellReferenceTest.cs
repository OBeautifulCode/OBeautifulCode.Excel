// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CellReferenceTest.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.Test
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics.CodeAnalysis;
    using System.Linq;

    using FakeItEasy;

    using FluentAssertions;

    using OBeautifulCode.CodeGen.ModelObject.Recipes;
    using OBeautifulCode.Excel.Test.Internal;

    using Xunit;

    public static partial class CellReferenceTest
    {
        [SuppressMessage("Microsoft.Performance", "CA1810:InitializeReferenceTypeStaticFieldsInline", Justification = ObcSuppressBecause.CA1810_InitializeReferenceTypeStaticFieldsInline_FieldsDeclaredInCodeGeneratedPartialTestClass)]
        static CellReferenceTest()
        {
            StringRepresentationTestScenarios
                .RemoveAllScenarios()
                .AddScenario(
                    new StringRepresentationTestScenario<CellReference>
                    {
                        Name = "ToString() should return 'KNOWN MISSING' when the cell is Known Missing.",
                        SystemUnderTestExpectedStringRepresentationFunc = () =>
                            new SystemUnderTestExpectedStringRepresentation<CellReference>
                            {
                                SystemUnderTest = CellReference.GetKnownMissing(),
                                ExpectedStringRepresentation = "KNOWN MISSING",
                            },
                    })
                .AddScenario(
                    new StringRepresentationTestScenario<CellReference>
                    {
                        Name = "ToString() should return WorksheetQualifiedA1Reference when cell is NOT Known Missing.",
                        SystemUnderTestExpectedStringRepresentationFunc = () =>
                        {
                            var systemUnderTest = A.Dummy<CellReference>();

                            var result = new SystemUnderTestExpectedStringRepresentation<CellReference>
                            {
                                SystemUnderTest = systemUnderTest,
                                ExpectedStringRepresentation = systemUnderTest.WorksheetQualifiedA1Reference,
                            };

                            return result;
                        },
                    });
        }

        [Fact]
        public static void Constructor___Should_throw_ArgumentException___When_parameter_worksheetName_is_malformed()
        {
            // Arrange
            var worksheetNames = new[]
            {
                @"\",
                @"/",
                @"*",
                @"[",
                @"]",
                @":",
                @"?",
                @"a\a",
                @"a/a",
                @"a*a",
                @"a[a",
                @"a]a",
                @"a:a",
                @"a?a",
                @"'",
                @"a'",
                @"'a",
                @"''",
                @"'aa",
                @"aa'",
                @"'aa'",
                @"ABCDEFGHIJKLMNOPQRSTUVWXYZ123456",
                @"ABCDEFGHIJKLMNOPQRSTUVWXYZ1234'",
                @"'BCDEFGHIJKLMNOPQRSTUVWXYZ12345",
                @"'BCDEFGHIJKLMNOPQRSTUVWXYZ1234'",
                @"ABCDEFGHIJKLMNOPQRSTUVWXYZ234'",
                @"'BCDEFGHIJKLMNOPQRSTUVWXYZ2345",
                @"'BCDEFGHIJKLMNOPQRSTUVWXYZ234'",
                "abc\r\ndef",
                "abc\rdef",
                "abc\ndef",
            };

            // Act
            var actuals = worksheetNames.Select(_ => Record.Exception(() => new CellReference(_, 1, 1))).ToList();

            // Assert
            foreach (var actual in actuals)
            {
                actual.Should().BeOfType<ArgumentException>();
                actual.Message.Should().Contain("worksheetName");
                actual.Message.Should().Contain("specified regex");
                actual.Message.Should().Contain("Worksheet names must have >= 1 and <= 31 characters");
            }
        }

        [Fact]
        public static void Constructor___Should_throw_ArgumentOutOfRangeException___When_parameter_rowNumber_is_less_than_1()
        {
            // Arrange
            var worksheetName = "name";
            var rowNumbers = new[] { 0, -1, int.MinValue };

            // Act
            var actuals = rowNumbers.Select(_ => Record.Exception(() => new CellReference(worksheetName, _, 1))).ToList();

            // Assert
            foreach (var actual in actuals)
            {
                actual.Should().BeOfType<ArgumentOutOfRangeException>();
                actual.Message.Should().Contain("rowNumber");
            }
        }

        [Fact]
        public static void Constructor___Should_throw_ArgumentOutOfRangeException___When_parameter_rowNumber_is_greater_than_1048576()
        {
            // Arrange
            var worksheetName = "name";
            var rowNumbers = new[] { 1048577, int.MaxValue };

            // Act
            var actuals = rowNumbers.Select(_ => Record.Exception(() => new CellReference(worksheetName, _, 1))).ToList();

            // Assert
            foreach (var actual in actuals)
            {
                actual.Should().BeOfType<ArgumentOutOfRangeException>();
                actual.Message.Should().Contain("rowNumber");
            }
        }

        [Fact]
        public static void Constructor___Should_throw_ArgumentOutOfRangeException___When_parameter_columnNumber_is_less_than_1()
        {
            // Arrange
            var worksheetName = "name";
            var columnNumbers = new[] { 0, -1, int.MinValue };

            // Act
            var actuals = columnNumbers.Select(_ => Record.Exception(() => new CellReference(worksheetName, 1, _))).ToList();

            // Assert
            foreach (var actual in actuals)
            {
                actual.Should().BeOfType<ArgumentOutOfRangeException>();
                actual.Message.Should().Contain("columnNumber");
            }
        }

        [Fact]
        public static void Constructor___Should_throw_ArgumentOutOfRangeException___When_parameter_columnNumber_is_greater_than_16384()
        {
            // Arrange
            var worksheetName = "name";
            var columnNumbers = new[] { 16385, int.MaxValue };

            // Act
            var actuals = columnNumbers.Select(_ => Record.Exception(() => new CellReference(worksheetName, 1, _))).ToList();

            // Assert
            foreach (var actual in actuals)
            {
                actual.Should().BeOfType<ArgumentOutOfRangeException>();
                actual.Message.Should().Contain("columnNumber");
            }
        }

        [Fact]
        public static void Constructor___Should_not_throw___When_all_parameters_are_valid()
        {
            // Arrange
            var worksheetNames = new[]
            {
                @"a",
                @"a'a",
                @"ABCDEFGHIJKLMNOPQRSTUVWXYZabcde",
                @"fghijklmnopqrstuvwxyz",
                @"1234567890",
                @" !""#$%&'()+,-.;<=>@^_`{|}~",
            };

            var rowNumbers = new[] { 1, 1000, 1048576 };
            var columnNumbers = new[] { 1, 1000, 16384 };

            // Act
            var actuals = new List<Exception>();
            foreach (var worksheetName in worksheetNames)
            {
                foreach (var rowNumber in rowNumbers)
                {
                    foreach (var columnNumber in columnNumbers)
                    {
                        actuals.Add(Record.Exception(() => new CellReference(worksheetName, rowNumber, columnNumber)));
                    }
                }
            }

            // Assert
            foreach (var actual in actuals)
            {
                actual.Should().BeNull();
            }
        }

        [Fact]
        public static void A1Reference___Should_return_the_A1_reference_to_the_cell___When_called()
        {
            // Arrange
            var worksheetName = "my-worksheet";

            var rowNumber = 569484;
            var columnNumber = 904;

            var expected = "AHT569484";

            var systemUnderTest = new CellReference(worksheetName, rowNumber, columnNumber);

            // Act
            var actual = systemUnderTest.A1Reference;

            // Assert
            actual.Should().Be(expected);
        }

        [Fact]
        public static void WorksheetQualifiedA1Reference___Should_return_the_worksheet_qualified_A1_reference_to_the_cell___When_called()
        {
            // Arrange
            var worksheetName1 = "my-worksheet";
            var worksheetName2 = "my'work'sheet";

            var rowNumber = 569484;
            var columnNumber = 904;

            var expected1 = "'my-worksheet'!AHT569484";
            var expected2 = "'my''work''sheet'!AHT569484";

            var systemUnderTest1 = new CellReference(worksheetName1, rowNumber, columnNumber);
            var systemUnderTest2 = new CellReference(worksheetName2, rowNumber, columnNumber);

            // Act
            var actual1 = systemUnderTest1.WorksheetQualifiedA1Reference;
            var actual2 = systemUnderTest2.WorksheetQualifiedA1Reference;

            // Assert
            actual1.Should().Be(expected1);
            actual2.Should().Be(expected2);
        }

        [Fact]
        public static void GetKnownMissing___Should_return_a_missing_cell_reference___When_called()
        {
            // Arrange
            var expected = new CellReference(@" !""#$%&'()+,-.;<=>@^_`{|}~54320", 1, 1);

            // Act
            var actual = CellReference.GetKnownMissing();

            // Assert
            actual.Should().Be(expected);
        }

        [Fact]
        public static void IsKnownMissing___Should_return_true___When_cellReference_indicates_a_missing_cell()
        {
            // Arrange
            var subjectUnderTest = CellReference.GetKnownMissing();

            // Act
            var actual = subjectUnderTest.IsKnownMissing();

            // Assert
            actual.Should().BeTrue();
        }

        [Fact]
        public static void IsKnownMissing___Should_return_false___When_cellReference_does_not_indicate_that_the_cell_is_missing()
        {
            // Arrange
            var subjectUnderTest = A.Dummy<CellReference>();

            // Act
            var actual = subjectUnderTest.IsKnownMissing();

            // Assert
            actual.Should().BeFalse();
        }

        [Fact]
        public static void FromA1Reference___Should_throw_ArgumentNullException___When_a1Reference_is_null()
        {
            // Arrange
            var worksheetName = "worksheet";

            // Act
            var actual = Record.Exception(() => CellReference.FromA1Reference(worksheetName, null));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("a1Reference");
        }

        [Fact]
        public static void FromA1Reference___Should_throw_ArgumentException___When_a1Reference_is_white_space()
        {
            // Arrange
            var worksheetName = "worksheet";

            // Act
            var actual = Record.Exception(() => CellReference.FromA1Reference(worksheetName, "  \r\n  "));

            // Assert
            actual.Should().BeOfType<ArgumentException>();
            actual.Message.Should().Contain("a1Reference");
            actual.Message.Should().Contain("white space");
        }

        [Fact]
        public static void FromA1Reference___Should_throw_ArgumentException___When_a1Reference_is_invalid()
        {
            // Arrange
            var worksheetName = "worksheet";
            var a1References = new[] { "A", "5", " A5 ", "AAAA3", "A11111111", "*", "5A", "A5A", "5A5" };

            // Act
            var actuals = a1References.Select(_ => Record.Exception(() => CellReference.FromA1Reference(worksheetName, _))).ToList();

            // Assert
            foreach (var actual in actuals)
            {
                actual.Should().BeOfType<ArgumentException>();
                actual.Message.Should().Contain("a1Reference");
                actual.Message.Should().Contain("is not matched by the specified regex");
            }
        }

        [Fact]
        public static void FromA1Reference___Should_throw_ArgumentOutOfRangeException___When_parsed_column_number_is_greater_than_MaximumColumnNumber()
        {
            // Arrange
            var worksheetName = "worksheet";
            var a1References = new[] { "XFE1", "ZZZ1" };

            // Act
            var actuals = a1References.Select(_ => Record.Exception(() => CellReference.FromA1Reference(worksheetName, _))).ToList();

            // Assert
            foreach (var actual in actuals)
            {
                actual.Should().BeOfType<ArgumentOutOfRangeException>();
                actual.Message.Should().Contain("columnNumber");
                actual.Message.Should().Contain(Constants.MaximumColumnNumber.ToString());
            }
        }

        [Fact]
        public static void FromA1Reference___Should_throw_ArgumentOutOfRangeException___When_parsed_row_number_is_greater_than_MaximumRowNumber()
        {
            // Arrange
            var worksheetName = "worksheet";
            var a1References = new[] { "A1048577", "A9999999" };

            // Act
            var actuals = a1References.Select(_ => Record.Exception(() => CellReference.FromA1Reference(worksheetName, _))).ToList();

            // Assert
            foreach (var actual in actuals)
            {
                actual.Should().BeOfType<ArgumentOutOfRangeException>();
                actual.Message.Should().Contain("rowNumber");
                actual.Message.Should().Contain(Constants.MaximumRowNumber.ToString());
            }
        }

        [Fact]
        public static void FromA1Reference___Should_return_CellReference_corresponding_to_specified_a1Reference___When_called()
        {
            var worksheetName = "worksheet-234234";

            var a1ReferenceToExpectedCellReferenceMap = new Dictionary<string, CellReference>
            {
                { "A1", new CellReference(worksheetName, 1, 1) },
                { "B1", new CellReference(worksheetName, 1, 2) },
                { "A2", new CellReference(worksheetName, 2, 1) },
                { "Z9", new CellReference(worksheetName, 9, 26) },
                { "Z99", new CellReference(worksheetName, 99, 26) },
                { "AA1", new CellReference(worksheetName, 1, 27) },
                { "AZ423", new CellReference(worksheetName, 423, 52) },
                { "BA99237", new CellReference(worksheetName, 99237, 53) },
                { "ZY2992", new CellReference(worksheetName, 2992, 701) },
                { "ZZ1048576", new CellReference(worksheetName, 1048576, 702) },
                { "AAA1048576", new CellReference(worksheetName, 1048576, 703) },
                { "AAB1048576", new CellReference(worksheetName, 1048576, 704) },
                { "OGR1048576", new CellReference(worksheetName, 1048576, 10340) },
                { "XFD1048576", new CellReference(worksheetName, 1048576, 16384) },
            };

            var expected = a1ReferenceToExpectedCellReferenceMap.OrderBy(_ => _.Key).Select(_ => _.Value);

            // Act
            var actual = a1ReferenceToExpectedCellReferenceMap.OrderBy(_ => _.Key).Select(_ => CellReference.FromA1Reference(worksheetName, _.Key)).ToList();

            // Assert
            expected.Should().Equal(actual);
        }

        [Fact]
        public static void FromWorksheetQualifiedA1Reference___Should_throw_ArgumentNullException___When_worksheetQualifiedA1Reference_is_null()
        {
            // Arrange, Act
            var actual = Record.Exception(() => CellReference.FromWorksheetQualifiedA1Reference(null));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("worksheetQualifiedA1Reference");
        }

        [Fact]
        public static void FromWorksheetQualifiedA1Reference___Should_throw_ArgumentException___When_worksheetQualifiedA1Reference_is_white_space()
        {
            // Arrange, Act
            var actual = Record.Exception(() => CellReference.FromWorksheetQualifiedA1Reference("  \r\n  "));

            // Assert
            actual.Should().BeOfType<ArgumentException>();
            actual.Message.Should().Contain("worksheetQualifiedA1Reference");
            actual.Message.Should().Contain("white space");
        }

        [Fact]
        public static void FromWorksheetQualifiedA1Reference___Should_throw_ArgumentException___When_worksheetQualifiedA1Reference_does_not_contain_exclamation_point()
        {
            // Arrange
            var worksheetQualifiedA1References = new[] { "A1", "Worksheet", "WorksheetA1" };

            // Act
            var actuals = worksheetQualifiedA1References.Select(_ => Record.Exception(() => CellReference.FromWorksheetQualifiedA1Reference(_))).ToList();

            // Assert
            foreach (var actual in actuals)
            {
                actual.Should().BeOfType<ArgumentException>();
                actual.Message.Should().Contain("Provided value (name: 'worksheetQualifiedA1Reference') does not contain the specified comparison value.");
                actual.Message.Should().Contain("Specified 'comparisonValue' is '!'");
            }
        }

        [Fact]
        public static void FromWorksheetQualifiedA1Reference___Should_throw_ArgumentException___When_worksheetQualifiedA1Reference_contains_invalid_worksheet_name()
        {
            // Arrange
            var worksheetQualifiedA1References = new[] { "!A1", "'!A1", "''!A1", "'''!A1", "''''!A1", "?A1", "'?'!A1" };

            // Act
            var actuals = worksheetQualifiedA1References.Select(_ => Record.Exception(() => CellReference.FromWorksheetQualifiedA1Reference(_))).ToList();

            // Assert
            foreach (var actual in actuals)
            {
                (actual is ArgumentException).Should().BeTrue();
            }
        }

        [Fact]
        public static void FromWorksheetQualifiedA1Reference___Should_throw_ArgumentException___When_worksheetQualifiedA1Reference_contains_worksheet_name_not_surrounded_with_single_quotes()
        {
            // Arrange
            var worksheetQualifiedA1References = new[] { "worksheet'!a1", "'worksheet!a1", "worksheet!A1" };

            // Act
            var actuals = worksheetQualifiedA1References.Select(_ => Record.Exception(() => CellReference.FromWorksheetQualifiedA1Reference(_))).ToList();

            // Assert
            foreach (var actual in actuals)
            {
                actual.Should().BeOfType<ArgumentException>();
            }
        }

        [Fact]
        public static void FromWorksheetQualifiedA1Reference___Should_throw_ArgumentException___When_worksheetQualifiedA1Reference_contains_invalid_a1Reference()
        {
            // Arrange
            var worksheetQualifiedA1References = new[] { "'worksheet'!A", "'worksheet'!1", "'worksheet'!A1A" };

            // Act
            var actuals = worksheetQualifiedA1References.Select(_ => Record.Exception(() => CellReference.FromWorksheetQualifiedA1Reference(_))).ToList();

            // Assert
            foreach (var actual in actuals)
            {
                actual.Should().BeOfType<ArgumentException>();
            }
        }

        [Fact]
        public static void FromWorksheetQualifiedA1Reference___Should_return_CellReference_corresponding_to_specified_worksheetQualifiedA1Reference___When_called()
        {
            var a1ReferenceToExpectedCellReferenceMap = new Dictionary<string, CellReference>
            {
                { "'a'!a1", new CellReference("a", 1, 1) },
                { "'&'!a1", new CellReference("&", 1, 1) },
                { "'my worksheet'!a1", new CellReference("my worksheet", 1, 1) },
                { "'my worksheet'!A1", new CellReference("my worksheet", 1, 1) },
                { "'my worksheet'!B1", new CellReference("my worksheet", 1, 2) },
                { "'my worksheet'!A2", new CellReference("my worksheet", 2, 1) },
                { "'my worksheet'!Z9", new CellReference("my worksheet", 9, 26) },
                { "'my worksheet'!Z99", new CellReference("my worksheet", 99, 26) },
                { "'my worksheet'!AA1", new CellReference("my worksheet", 1, 27) },
                { "'my worksheet'!AZ423", new CellReference("my worksheet", 423, 52) },
                { "'my worksheet'!BA99237", new CellReference("my worksheet", 99237, 53) },
                { "'my worksheet'!ZY2992", new CellReference("my worksheet", 2992, 701) },
                { "'my worksheet'!ZZ1048576", new CellReference("my worksheet", 1048576, 702) },
                { "'my worksheet'!AAA1048576", new CellReference("my worksheet", 1048576, 703) },
                { "'my worksheet'!AAB1048576", new CellReference("my worksheet", 1048576, 704) },
                { "'my worksheet'!OGR1048576", new CellReference("my worksheet", 1048576, 10340) },
                { "'my worksheet'!XFD1048576", new CellReference("my worksheet", 1048576, 16384) },
            };

            var expected = a1ReferenceToExpectedCellReferenceMap.OrderBy(_ => _.Key).Select(_ => _.Value);

            // Act
            var actual = a1ReferenceToExpectedCellReferenceMap.OrderBy(_ => _.Key).Select(_ => CellReference.FromWorksheetQualifiedA1Reference(_.Key)).ToList();

            // Assert
            expected.Should().Equal(actual);
        }
    }
}
