// --------------------------------------------------------------------------------------------------------------------
// <copyright file="NamedCellTest.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.Test
{
    using System;
    using System.Linq;

    using FakeItEasy;

    using FluentAssertions;

    using Naos.Serialization.Bson;
    using Naos.Serialization.Json;

    using OBeautifulCode.AutoFakeItEasy;
    using OBeautifulCode.Excel.Serialization.Bson;

    using Xunit;

    public static class NamedCellTest
    {
        private static readonly NamedCell ObjectForEquatableTests = A.Dummy<NamedCell>();

        private static readonly NamedCell ObjectThatIsEqualButNotTheSameAsObjectForEquatableTests =
            new NamedCell(ObjectForEquatableTests.Name, ObjectForEquatableTests.Cell);

        private static readonly NamedCell[] ObjectsThatAreNotEqualToObjectForEquatableTests =
        {
            A.Dummy<NamedCell>(),
            new NamedCell(A.Dummy<string>(), ObjectForEquatableTests.Cell),
            new NamedCell(ObjectForEquatableTests.Name, A.Dummy<CellReference>().ThatIsNot(ObjectForEquatableTests.Cell)),
        };

        private static readonly string ObjectThatIsNotTheSameTypeAsObjectForEquatableTests = A.Dummy<string>();

        [Fact]
        public static void Constructor___Should_throw_ArgumentNullException___When_parameter_name_is_null()
        {
            // Arrange
            var cell = A.Dummy<CellReference>();

            // Act
            var actual = Record.Exception(() => new NamedCell(null, cell));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("name");
        }

        [Fact]
        public static void Constructor___Should_throw_ArgumentException___When_parameter_name_is_white_space()
        {
            // Arrange
            var cell = A.Dummy<CellReference>();

            // Act
            var actual = Record.Exception(() => new NamedCell(" \r\n ", cell));

            // Assert
            actual.Should().BeOfType<ArgumentException>();
            actual.Message.Should().Contain("name");
            actual.Message.Should().Contain("white space");
        }

        [Fact]
        public static void Constructor___Should_throw_ArgumentNullException___When_parameter_cell_is_null()
        {
            // Arrange
            var name = A.Dummy<string>();

            // Act
            var actual = Record.Exception(() => new NamedCell(name, null));

            // Assert
            actual.Should().BeOfType<ArgumentNullException>();
            actual.Message.Should().Contain("cell");
        }

        [Fact]
        public static void Name___Should_return_same_name_passed_to_constructor___When_called()
        {
            // Arrange
            var expected = A.Dummy<string>();
            var cell = A.Dummy<CellReference>();
            var systemUnderTest = new NamedCell(expected, cell);

            // Act
            var actual = systemUnderTest.Name;

            // Assert
            actual.Should().Be(expected);
        }

        [Fact]
        public static void Cell___Should_return_same_cell_passed_to_constructor___When_called()
        {
            // Arrange
            var name = A.Dummy<string>();
            var expected = A.Dummy<CellReference>();
            var systemUnderTest = new NamedCell(name, expected);

            // Act
            var actual = systemUnderTest.Cell;

            // Assert
            actual.Should().Be(expected);
        }

        [Fact]
        public static void Deserialize___Should_roundtrip_object___When_serializing_and_deserializing_using_NaosJsonSerializer()
        {
            // Arrange
            var expected = A.Dummy<NamedCell>();
            var serializer = new NaosJsonSerializer();
            var serializedJson = serializer.SerializeToString(expected);

            // Act
            var actual = serializer.Deserialize<NamedCell>(serializedJson);

            // Assert
            actual.Should().Be(expected);
        }

        [Fact]
        public static void Deserialize___Should_roundtrip_object___When_serializing_and_deserializing_using_NaosBsonSerializer()
        {
            // Arrange
            var expected = A.Dummy<NamedCell>();
            var serializer = new NaosBsonSerializer(configurationType: typeof(ExcelBsonConfiguration));

            var serializedBson = serializer.SerializeToString(expected);

            // Act
            var actual = serializer.Deserialize<NamedCell>(serializedBson);

            // Assert
            actual.Should().Be(expected);
        }

        [Fact]
        public static void EqualsOperator___Should_return_true___When_both_sides_of_operator_are_null()
        {
            // Arrange
            NamedCell systemUnderTest1 = null;
            NamedCell systemUnderTest2 = null;

            // Act
            var result = systemUnderTest1 == systemUnderTest2;

            // Assert
            result.Should().BeTrue();
        }

        [Fact]
        public static void EqualsOperator___Should_return_false___When_one_side_of_operator_is_null_and_the_other_side_is_not_null()
        {
            // Arrange
            NamedCell systemUnderTest = null;

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
            NamedCell systemUnderTest1 = null;
            NamedCell systemUnderTest2 = null;

            // Act
            var result = systemUnderTest1 != systemUnderTest2;

            // Assert
            result.Should().BeFalse();
        }

        [Fact]
        public static void NotEqualsOperator___Should_return_true___When_one_side_of_operator_is_null_and_the_other_side_is_not_null()
        {
            // Arrange
            NamedCell systemUnderTest = null;

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
        public static void Equals_with_NamedCell___Should_return_false___When_parameter_other_is_null()
        {
            // Arrange
            NamedCell systemUnderTest = null;

            // Act
            var result = ObjectForEquatableTests.Equals(systemUnderTest);

            // Assert
            result.Should().BeFalse();
        }

        [Fact]
        public static void Equals_with_NamedCell___Should_return_true___When_parameter_other_is_same_object()
        {
            // Arrange, Act
            var result = ObjectForEquatableTests.Equals(ObjectForEquatableTests);

            // Assert
            result.Should().BeTrue();
        }

        [Fact]
        public static void Equals_with_NamedCell___Should_return_false___When_objects_being_compared_have_different_property_values()
        {
            // Arrange, Act
            var results = ObjectsThatAreNotEqualToObjectForEquatableTests.Select(_ => ObjectForEquatableTests.Equals(_)).ToList();

            // Assert
            results.ForEach(_ => _.Should().BeFalse());
        }

        [Fact]
        public static void Equals_with_NamedCell___Should_return_true___When_objects_being_compared_have_same_property_values()
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

        [Fact]
        public static void DeepClone___Should_clone_item___When_called()
        {
            // Arrange
            var systemUnderTest = A.Dummy<NamedCell>();

            // Act
            var actual = systemUnderTest.DeepClone();

            // Assert
            actual.Should().Be(systemUnderTest);
            actual.Should().NotBeSameAs(systemUnderTest);
            actual.Cell.Should().NotBeSameAs(systemUnderTest.Cell);
        }
    }
}
