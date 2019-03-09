﻿// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CellReferenceTest.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.Test
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    using FakeItEasy;

    using FluentAssertions;

    using Naos.Serialization.Bson;
    using Naos.Serialization.Json;

    using OBeautifulCode.AutoFakeItEasy;
    using OBeautifulCode.Excel.Serialization.Bson;

    using Xunit;

    public static class CellReferenceTest
    {
        private static readonly CellReference ObjectForEquatableTests = A.Dummy<CellReference>();

        private static readonly CellReference ObjectThatIsEqualButNotTheSameAsObjectForEquatableTests =
            new CellReference(ObjectForEquatableTests.WorksheetName, ObjectForEquatableTests.RowNumber, ObjectForEquatableTests.ColumnNumber);

        private static readonly CellReference[] ObjectsThatAreNotEqualToObjectForEquatableTests =
        {
            A.Dummy<CellReference>(),
            new CellReference("worksheet-" + A.Dummy<Guid>().ToString().Substring(1, 10), ObjectForEquatableTests.RowNumber, ObjectForEquatableTests.ColumnNumber),
            new CellReference(ObjectForEquatableTests.WorksheetName, A.Dummy<PositiveInteger>().ThatIs(_ => (_ <= Constants.MaximumRowNumber) && (_ != ObjectForEquatableTests.RowNumber)), ObjectForEquatableTests.ColumnNumber),
            new CellReference(ObjectForEquatableTests.WorksheetName, ObjectForEquatableTests.RowNumber, A.Dummy<PositiveInteger>().ThatIs(_ => (_ <= Constants.MaximumColumnNumber) && (_ != ObjectForEquatableTests.ColumnNumber))),
        };

        private static readonly string ObjectThatIsNotTheSameTypeAsObjectForEquatableTests = A.Dummy<string>();

        [Fact]
        public static void Constructor___Should_throw_ArgumentNullException___When_parameter_worksheetName_is_null()
        {
            // Arrange, Act
            var actual = Record.Exception(() => new CellReference(null, 1, 1));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("worksheetName");
        }

        [Fact]
        public static void Constructor___Should_throw_ArgumentException___When_parameter_worksheetName_is_white_space()
        {
            // Arrange, Act
            var actual = Record.Exception(() => new CellReference(" \r\n  ", 1, 1));

            // Assert
            actual.Should().BeOfType<ArgumentException>();
            actual.Message.Should().Contain("worksheetName");
            actual.Message.Should().Contain("white space");
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
        public static void WorksheetName___Should_return_same_worksheetName_passed_to_constructor___When_called()
        {
            // Arrange
            var expected = "my-worksheet";
            var systemUnderTest = new CellReference(expected, 1, 1);

            // Act
            var actual = systemUnderTest.WorksheetName;

            // Assert
            actual.Should().Be(expected);
        }

        [Fact]
        public static void RowNumber___Should_return_same_rowNumber_passed_to_constructor___When_called()
        {
            // Arrange
            var worksheetName = "my-worksheet";
            var expected = A.Dummy<PositiveInteger>().ThatIs(_ => _ <= Constants.MaximumRowNumber);
            var systemUnderTest = new CellReference(worksheetName, expected, 1);

            // Act
            var actual = systemUnderTest.RowNumber;

            // Assert
            actual.Should().Be(expected);
        }

        [Fact]
        public static void ColumnNumber___Should_return_same_columnNumber_passed_to_constructor___When_called()
        {
            // Arrange
            var worksheetName = "my-worksheet";
            var expected = A.Dummy<PositiveInteger>().ThatIs(_ => _ <= Constants.MaximumColumnNumber);
            var systemUnderTest = new CellReference(worksheetName, 1, expected);

            // Act
            var actual = systemUnderTest.ColumnNumber;

            // Assert
            actual.Should().Be(expected);
        }

        [Fact]
        public static void WorksheetQualifiedA1Reference___Should_return_the_worksheet_qualified_cell_name___When_called()
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
        public static void Deserialize___Should_roundtrip_object___When_serializing_and_deserializing_using_NaosJsonSerializer()
        {
            // Arrange
            var expected = A.Dummy<CellReference>();
            var serializer = new NaosJsonSerializer();
            var serializedJson = serializer.SerializeToString(expected);

            // Act
            var actual = serializer.Deserialize<CellReference>(serializedJson);

            // Assert
            actual.Should().Be(expected);
        }

        [Fact]
        public static void Deserialize___Should_roundtrip_object___When_serializing_and_deserializing_using_NaosBsonSerializer()
        {
            // Arrange
            var expected = A.Dummy<CellReference>();
            var serializer = new NaosBsonSerializer(configurationType: typeof(ExcelBsonConfiguration));

            var serializedBson = serializer.SerializeToString(expected);

            // Act
            var actual = serializer.Deserialize<CellReference>(serializedBson);

            // Assert
            actual.Should().Be(expected);
        }

        [Fact]
        public static void EqualsOperator___Should_return_true___When_both_sides_of_operator_are_null()
        {
            // Arrange
            CellReference systemUnderTest1 = null;
            CellReference systemUnderTest2 = null;

            // Act
            var result = systemUnderTest1 == systemUnderTest2;

            // Assert
            result.Should().BeTrue();
        }

        [Fact]
        public static void EqualsOperator___Should_return_false___When_one_side_of_operator_is_null_and_the_other_side_is_not_null()
        {
            // Arrange
            CellReference systemUnderTest = null;

            // Act
            var result1 = systemUnderTest == ObjectForEquatableTests;
            var result2 = ObjectForEquatableTests == systemUnderTest;

            // Assert
            result1.Should().BeFalse();
            result2.Should().BeFalse();
        }

        [Fact]
        public static void EqualsOperator___Should_return_true___When_same_object_is_on_both_sides_of_operator()
        {
            // Arrange, Act
#pragma warning disable CS1718 // Comparison made to same variable
            var result = ObjectForEquatableTests == ObjectForEquatableTests;
#pragma warning restore CS1718 // Comparison made to same variable

            // Assert
            result.Should().BeTrue();
        }

        [Fact]
        public static void EqualsOperator___Should_return_false___When_objects_being_compared_have_different_property_values()
        {
            // Arrange, Act
            var results = ObjectsThatAreNotEqualToObjectForEquatableTests.Select(_ => ObjectForEquatableTests == _).ToList();

            // Assert
            results.ForEach(_ => _.Should().BeFalse());
        }

        [Fact]
        public static void EqualsOperator___Should_return_true___When_objects_being_compared_have_same_property_values()
        {
            // Arrange, Act
            var result = ObjectForEquatableTests == ObjectThatIsEqualButNotTheSameAsObjectForEquatableTests;

            // Assert
            result.Should().BeTrue();
        }

        [Fact]
        public static void NotEqualsOperator___Should_return_false___When_both_sides_of_operator_are_null()
        {
            // Arrange
            CellReference systemUnderTest1 = null;
            CellReference systemUnderTest2 = null;

            // Act
            var result = systemUnderTest1 != systemUnderTest2;

            // Assert
            result.Should().BeFalse();
        }

        [Fact]
        public static void NotEqualsOperator___Should_return_true___When_one_side_of_operator_is_null_and_the_other_side_is_not_null()
        {
            // Arrange
            CellReference systemUnderTest = null;

            // Act
            var result1 = systemUnderTest != ObjectForEquatableTests;
            var result2 = ObjectForEquatableTests != systemUnderTest;

            // Assert
            result1.Should().BeTrue();
            result2.Should().BeTrue();
        }

        [Fact]
        public static void NotEqualsOperator___Should_return_false___When_same_object_is_on_both_sides_of_operator()
        {
            // Arrange, Act
#pragma warning disable CS1718 // Comparison made to same variable
            var result = ObjectForEquatableTests != ObjectForEquatableTests;
#pragma warning restore CS1718 // Comparison made to same variable

            // Assert
            result.Should().BeFalse();
        }

        [Fact]
        public static void NotEqualsOperator___Should_return_true___When_objects_being_compared_have_different_property_values()
        {
            // Arrange, Act
            var results = ObjectsThatAreNotEqualToObjectForEquatableTests.Select(_ => ObjectForEquatableTests != _).ToList();

            // sAssert
            results.ForEach(_ => _.Should().BeTrue());
        }

        [Fact]
        public static void NotEqualsOperator___Should_return_false___When_objects_being_compared_have_same_property_values()
        {
            // Arrange, Act
            var result = ObjectForEquatableTests != ObjectThatIsEqualButNotTheSameAsObjectForEquatableTests;

            // Assert
            result.Should().BeFalse();
        }

        [Fact]
        public static void Equals_with_CellReference___Should_return_false___When_parameter_other_is_null()
        {
            // Arrange
            CellReference systemUnderTest = null;

            // Act
            var result = ObjectForEquatableTests.Equals(systemUnderTest);

            // Assert
            result.Should().BeFalse();
        }

        [Fact]
        public static void Equals_with_CellReference___Should_return_true___When_parameter_other_is_same_object()
        {
            // Arrange, Act
            var result = ObjectForEquatableTests.Equals(ObjectForEquatableTests);

            // Assert
            result.Should().BeTrue();
        }

        [Fact]
        public static void Equals_with_CellReference___Should_return_false___When_objects_being_compared_have_different_property_values()
        {
            // Arrange, Act
            var results = ObjectsThatAreNotEqualToObjectForEquatableTests.Select(_ => ObjectForEquatableTests.Equals(_)).ToList();

            // Assert
            results.ForEach(_ => _.Should().BeFalse());
        }

        [Fact]
        public static void Equals_with_CellReference___Should_return_true___When_objects_being_compared_have_same_property_values()
        {
            // Arrange, Act
            var result = ObjectForEquatableTests.Equals(ObjectThatIsEqualButNotTheSameAsObjectForEquatableTests);

            // Assert
            result.Should().BeTrue();
        }

        [Fact]
        public static void Equals_with_Object___Should_return_false___When_parameter_other_is_null()
        {
            // Arrange, Act
            var result = ObjectForEquatableTests.Equals(null);

            // Assert
            result.Should().BeFalse();
        }

        [Fact]
        public static void Equals_with_Object___Should_return_false___When_parameter_other_is_not_of_the_same_type()
        {
            // Arrange, Act
            var result = ObjectForEquatableTests.Equals((object)ObjectThatIsNotTheSameTypeAsObjectForEquatableTests);

            // Assert
            result.Should().BeFalse();
        }

        [Fact]
        public static void Equals_with_Object___Should_return_true___When_parameter_other_is_same_object()
        {
            // Arrange, Act
            var result = ObjectForEquatableTests.Equals((object)ObjectForEquatableTests);

            // Assert
            result.Should().BeTrue();
        }

        [Fact]
        public static void Equals_with_Object___Should_return_false___When_objects_being_compared_have_different_property_values()
        {
            // Arrange, Act
            var results = ObjectsThatAreNotEqualToObjectForEquatableTests.Select(_ => ObjectForEquatableTests.Equals((object)_)).ToList();

            // Assert
            results.ForEach(_ => _.Should().BeFalse());
        }

        [Fact]
        public static void Equals_with_Object___Should_return_true___When_objects_being_compared_have_same_property_values()
        {
            // Arrange, Act
            var result = ObjectForEquatableTests.Equals((object)ObjectThatIsEqualButNotTheSameAsObjectForEquatableTests);

            // Assert
            result.Should().BeTrue();
        }

        [Fact]
        public static void GetHashCode___Should_not_be_equal_for_two_objects___When_objects_have_different_property_values()
        {
            // Arrange, Act
            var hashCode1 = ObjectForEquatableTests.GetHashCode();
            var hashCode2 = ObjectsThatAreNotEqualToObjectForEquatableTests.Select(_ => _.GetHashCode()).ToList();

            // Assert
            hashCode2.ForEach(_ => _.Should().NotBe(hashCode1));
        }

        [Fact]
        public static void GetHashCode___Should_be_equal_for_two_objects___When_objects_have_the_same_property_values()
        {
            // Arrange, Act
            var hash1 = ObjectForEquatableTests.GetHashCode();
            var hash2 = ObjectThatIsEqualButNotTheSameAsObjectForEquatableTests.GetHashCode();

            // Assert
            hash1.Should().Be(hash2);
        }
    }
}