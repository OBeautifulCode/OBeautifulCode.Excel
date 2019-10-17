// --------------------------------------------------------------------------------------------------------------------
// <copyright file="NumericDataValidationTest.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.Test
{
    using System.Linq;

    using FakeItEasy;

    using FluentAssertions;

    using OBeautifulCode.AutoFakeItEasy;
    using OBeautifulCode.Excel.Serialization.Bson;
    using OBeautifulCode.Excel.Serialization.Json;
    using OBeautifulCode.Serialization.Bson;
    using OBeautifulCode.Serialization.Json;

    using Xunit;

    public static class NumericDataValidationTest
    {
        private static readonly NumericDataValidation ObjectForEquatableTests = A.Dummy<NumericDataValidation>();

        private static readonly NumericDataValidation ObjectThatIsEqualButNotTheSameAsObjectForEquatableTests =
            new NumericDataValidation
            {
                Kind = ObjectForEquatableTests.Kind,
                Operator = ObjectForEquatableTests.Operator,
                Operand1Formula = ObjectForEquatableTests.Operand1Formula,
                Operand2Formula = ObjectForEquatableTests.Operand2Formula,
                IgnoreBlank = ObjectForEquatableTests.IgnoreBlank,
                ShowInputMessage = ObjectForEquatableTests.ShowInputMessage,
                InputMessageTitle = ObjectForEquatableTests.InputMessageTitle,
                InputMessageBody = ObjectForEquatableTests.InputMessageBody,
                ShowErrorAlertAfterInvalidDataIsEntered = ObjectForEquatableTests.ShowErrorAlertAfterInvalidDataIsEntered,
                ErrorAlertStyle = ObjectForEquatableTests.ErrorAlertStyle,
                ErrorAlertTitle = ObjectForEquatableTests.ErrorAlertTitle,
                ErrorAlertBody = ObjectForEquatableTests.ErrorAlertBody,
                ShowListDropdown = ObjectForEquatableTests.ShowListDropdown,
                Operand1NumericValue = ObjectForEquatableTests.Operand1NumericValue,
                Operand2NumericValue = ObjectForEquatableTests.Operand2NumericValue,
            };

        private static readonly NumericDataValidation[] ObjectsThatAreNotEqualToObjectForEquatableTests =
        {
            A.Dummy<NumericDataValidation>(),
            new NumericDataValidation(),
            new NumericDataValidation
            {
                Kind = A.Dummy<DataValidationKind>().ThatIsNot(ObjectForEquatableTests.Kind),
                Operator = ObjectForEquatableTests.Operator,
                Operand1Formula = ObjectForEquatableTests.Operand1Formula,
                Operand2Formula = ObjectForEquatableTests.Operand2Formula,
                IgnoreBlank = ObjectForEquatableTests.IgnoreBlank,
                ShowInputMessage = ObjectForEquatableTests.ShowInputMessage,
                InputMessageTitle = ObjectForEquatableTests.InputMessageTitle,
                InputMessageBody = ObjectForEquatableTests.InputMessageBody,
                ShowErrorAlertAfterInvalidDataIsEntered = ObjectForEquatableTests.ShowErrorAlertAfterInvalidDataIsEntered,
                ErrorAlertStyle = ObjectForEquatableTests.ErrorAlertStyle,
                ErrorAlertTitle = ObjectForEquatableTests.ErrorAlertTitle,
                ErrorAlertBody = ObjectForEquatableTests.ErrorAlertBody,
                ShowListDropdown = ObjectForEquatableTests.ShowListDropdown,
                Operand1NumericValue = ObjectForEquatableTests.Operand1NumericValue,
                Operand2NumericValue = ObjectForEquatableTests.Operand2NumericValue,
            },
            new NumericDataValidation
            {
                Kind = ObjectForEquatableTests.Kind,
                Operator = A.Dummy<DataValidationOperator>().ThatIsNot(ObjectForEquatableTests.Operator),
                Operand1Formula = ObjectForEquatableTests.Operand1Formula,
                Operand2Formula = ObjectForEquatableTests.Operand2Formula,
                IgnoreBlank = ObjectForEquatableTests.IgnoreBlank,
                ShowInputMessage = ObjectForEquatableTests.ShowInputMessage,
                InputMessageTitle = ObjectForEquatableTests.InputMessageTitle,
                InputMessageBody = ObjectForEquatableTests.InputMessageBody,
                ShowErrorAlertAfterInvalidDataIsEntered = ObjectForEquatableTests.ShowErrorAlertAfterInvalidDataIsEntered,
                ErrorAlertStyle = ObjectForEquatableTests.ErrorAlertStyle,
                ErrorAlertTitle = ObjectForEquatableTests.ErrorAlertTitle,
                ErrorAlertBody = ObjectForEquatableTests.ErrorAlertBody,
                ShowListDropdown = ObjectForEquatableTests.ShowListDropdown,
                Operand1NumericValue = ObjectForEquatableTests.Operand1NumericValue,
                Operand2NumericValue = ObjectForEquatableTests.Operand2NumericValue,
            },
            new NumericDataValidation
            {
                Kind = ObjectForEquatableTests.Kind,
                Operator = ObjectForEquatableTests.Operator,
                Operand1Formula = A.Dummy<string>(),
                Operand2Formula = ObjectForEquatableTests.Operand2Formula,
                IgnoreBlank = ObjectForEquatableTests.IgnoreBlank,
                ShowInputMessage = ObjectForEquatableTests.ShowInputMessage,
                InputMessageTitle = ObjectForEquatableTests.InputMessageTitle,
                InputMessageBody = ObjectForEquatableTests.InputMessageBody,
                ShowErrorAlertAfterInvalidDataIsEntered = ObjectForEquatableTests.ShowErrorAlertAfterInvalidDataIsEntered,
                ErrorAlertStyle = ObjectForEquatableTests.ErrorAlertStyle,
                ErrorAlertTitle = ObjectForEquatableTests.ErrorAlertTitle,
                ErrorAlertBody = ObjectForEquatableTests.ErrorAlertBody,
                ShowListDropdown = ObjectForEquatableTests.ShowListDropdown,
                Operand1NumericValue = ObjectForEquatableTests.Operand1NumericValue,
                Operand2NumericValue = ObjectForEquatableTests.Operand2NumericValue,
            },
            new NumericDataValidation
            {
                Kind = ObjectForEquatableTests.Kind,
                Operator = ObjectForEquatableTests.Operator,
                Operand1Formula = ObjectForEquatableTests.Operand1Formula,
                Operand2Formula = A.Dummy<string>(),
                IgnoreBlank = ObjectForEquatableTests.IgnoreBlank,
                ShowInputMessage = ObjectForEquatableTests.ShowInputMessage,
                InputMessageTitle = ObjectForEquatableTests.InputMessageTitle,
                InputMessageBody = ObjectForEquatableTests.InputMessageBody,
                ShowErrorAlertAfterInvalidDataIsEntered = ObjectForEquatableTests.ShowErrorAlertAfterInvalidDataIsEntered,
                ErrorAlertStyle = ObjectForEquatableTests.ErrorAlertStyle,
                ErrorAlertTitle = ObjectForEquatableTests.ErrorAlertTitle,
                ErrorAlertBody = ObjectForEquatableTests.ErrorAlertBody,
                ShowListDropdown = ObjectForEquatableTests.ShowListDropdown,
                Operand1NumericValue = ObjectForEquatableTests.Operand1NumericValue,
                Operand2NumericValue = ObjectForEquatableTests.Operand2NumericValue,
            },
            new NumericDataValidation
            {
                Kind = ObjectForEquatableTests.Kind,
                Operator = ObjectForEquatableTests.Operator,
                Operand1Formula = ObjectForEquatableTests.Operand1Formula,
                Operand2Formula = ObjectForEquatableTests.Operand2Formula,
                IgnoreBlank = !ObjectForEquatableTests.IgnoreBlank,
                ShowInputMessage = ObjectForEquatableTests.ShowInputMessage,
                InputMessageTitle = ObjectForEquatableTests.InputMessageTitle,
                InputMessageBody = ObjectForEquatableTests.InputMessageBody,
                ShowErrorAlertAfterInvalidDataIsEntered = ObjectForEquatableTests.ShowErrorAlertAfterInvalidDataIsEntered,
                ErrorAlertStyle = ObjectForEquatableTests.ErrorAlertStyle,
                ErrorAlertTitle = ObjectForEquatableTests.ErrorAlertTitle,
                ErrorAlertBody = ObjectForEquatableTests.ErrorAlertBody,
                ShowListDropdown = ObjectForEquatableTests.ShowListDropdown,
                Operand1NumericValue = ObjectForEquatableTests.Operand1NumericValue,
                Operand2NumericValue = ObjectForEquatableTests.Operand2NumericValue,
            },
            new NumericDataValidation
            {
                Kind = ObjectForEquatableTests.Kind,
                Operator = ObjectForEquatableTests.Operator,
                Operand1Formula = ObjectForEquatableTests.Operand1Formula,
                Operand2Formula = ObjectForEquatableTests.Operand2Formula,
                IgnoreBlank = ObjectForEquatableTests.IgnoreBlank,
                ShowInputMessage = !ObjectForEquatableTests.ShowInputMessage,
                InputMessageTitle = ObjectForEquatableTests.InputMessageTitle,
                InputMessageBody = ObjectForEquatableTests.InputMessageBody,
                ShowErrorAlertAfterInvalidDataIsEntered = ObjectForEquatableTests.ShowErrorAlertAfterInvalidDataIsEntered,
                ErrorAlertStyle = ObjectForEquatableTests.ErrorAlertStyle,
                ErrorAlertTitle = ObjectForEquatableTests.ErrorAlertTitle,
                ErrorAlertBody = ObjectForEquatableTests.ErrorAlertBody,
                ShowListDropdown = ObjectForEquatableTests.ShowListDropdown,
                Operand1NumericValue = ObjectForEquatableTests.Operand1NumericValue,
                Operand2NumericValue = ObjectForEquatableTests.Operand2NumericValue,
            },
            new NumericDataValidation
            {
                Kind = ObjectForEquatableTests.Kind,
                Operator = ObjectForEquatableTests.Operator,
                Operand1Formula = ObjectForEquatableTests.Operand1Formula,
                Operand2Formula = ObjectForEquatableTests.Operand2Formula,
                IgnoreBlank = ObjectForEquatableTests.IgnoreBlank,
                ShowInputMessage = ObjectForEquatableTests.ShowInputMessage,
                InputMessageTitle = A.Dummy<string>(),
                InputMessageBody = ObjectForEquatableTests.InputMessageBody,
                ShowErrorAlertAfterInvalidDataIsEntered = ObjectForEquatableTests.ShowErrorAlertAfterInvalidDataIsEntered,
                ErrorAlertStyle = ObjectForEquatableTests.ErrorAlertStyle,
                ErrorAlertTitle = ObjectForEquatableTests.ErrorAlertTitle,
                ErrorAlertBody = ObjectForEquatableTests.ErrorAlertBody,
                ShowListDropdown = ObjectForEquatableTests.ShowListDropdown,
                Operand1NumericValue = ObjectForEquatableTests.Operand1NumericValue,
                Operand2NumericValue = ObjectForEquatableTests.Operand2NumericValue,
            },
            new NumericDataValidation
            {
                Kind = ObjectForEquatableTests.Kind,
                Operator = ObjectForEquatableTests.Operator,
                Operand1Formula = ObjectForEquatableTests.Operand1Formula,
                Operand2Formula = ObjectForEquatableTests.Operand2Formula,
                IgnoreBlank = ObjectForEquatableTests.IgnoreBlank,
                ShowInputMessage = ObjectForEquatableTests.ShowInputMessage,
                InputMessageTitle = ObjectForEquatableTests.InputMessageTitle,
                InputMessageBody = A.Dummy<string>(),
                ShowErrorAlertAfterInvalidDataIsEntered = ObjectForEquatableTests.ShowErrorAlertAfterInvalidDataIsEntered,
                ErrorAlertStyle = ObjectForEquatableTests.ErrorAlertStyle,
                ErrorAlertTitle = ObjectForEquatableTests.ErrorAlertTitle,
                ErrorAlertBody = ObjectForEquatableTests.ErrorAlertBody,
                ShowListDropdown = ObjectForEquatableTests.ShowListDropdown,
                Operand1NumericValue = ObjectForEquatableTests.Operand1NumericValue,
                Operand2NumericValue = ObjectForEquatableTests.Operand2NumericValue,
            },
            new NumericDataValidation
            {
                Kind = ObjectForEquatableTests.Kind,
                Operator = ObjectForEquatableTests.Operator,
                Operand1Formula = ObjectForEquatableTests.Operand1Formula,
                Operand2Formula = ObjectForEquatableTests.Operand2Formula,
                IgnoreBlank = ObjectForEquatableTests.IgnoreBlank,
                ShowInputMessage = ObjectForEquatableTests.ShowInputMessage,
                InputMessageTitle = ObjectForEquatableTests.InputMessageTitle,
                InputMessageBody = ObjectForEquatableTests.InputMessageBody,
                ShowErrorAlertAfterInvalidDataIsEntered = !ObjectForEquatableTests.ShowErrorAlertAfterInvalidDataIsEntered,
                ErrorAlertStyle = ObjectForEquatableTests.ErrorAlertStyle,
                ErrorAlertTitle = ObjectForEquatableTests.ErrorAlertTitle,
                ErrorAlertBody = ObjectForEquatableTests.ErrorAlertBody,
                ShowListDropdown = ObjectForEquatableTests.ShowListDropdown,
                Operand1NumericValue = ObjectForEquatableTests.Operand1NumericValue,
                Operand2NumericValue = ObjectForEquatableTests.Operand2NumericValue,
            },
            new NumericDataValidation
            {
                Kind = ObjectForEquatableTests.Kind,
                Operator = ObjectForEquatableTests.Operator,
                Operand1Formula = ObjectForEquatableTests.Operand1Formula,
                Operand2Formula = ObjectForEquatableTests.Operand2Formula,
                IgnoreBlank = ObjectForEquatableTests.IgnoreBlank,
                ShowInputMessage = ObjectForEquatableTests.ShowInputMessage,
                InputMessageTitle = ObjectForEquatableTests.InputMessageTitle,
                InputMessageBody = ObjectForEquatableTests.InputMessageBody,
                ShowErrorAlertAfterInvalidDataIsEntered = ObjectForEquatableTests.ShowErrorAlertAfterInvalidDataIsEntered,
                ErrorAlertStyle = A.Dummy<DataValidationErrorAlertStyle>().ThatIsNot(ObjectForEquatableTests.ErrorAlertStyle),
                ErrorAlertTitle = ObjectForEquatableTests.ErrorAlertTitle,
                ErrorAlertBody = ObjectForEquatableTests.ErrorAlertBody,
                ShowListDropdown = ObjectForEquatableTests.ShowListDropdown,
                Operand1NumericValue = ObjectForEquatableTests.Operand1NumericValue,
                Operand2NumericValue = ObjectForEquatableTests.Operand2NumericValue,
            },
            new NumericDataValidation
            {
                Kind = ObjectForEquatableTests.Kind,
                Operator = ObjectForEquatableTests.Operator,
                Operand1Formula = ObjectForEquatableTests.Operand1Formula,
                Operand2Formula = ObjectForEquatableTests.Operand2Formula,
                IgnoreBlank = ObjectForEquatableTests.IgnoreBlank,
                ShowInputMessage = ObjectForEquatableTests.ShowInputMessage,
                InputMessageTitle = ObjectForEquatableTests.InputMessageTitle,
                InputMessageBody = ObjectForEquatableTests.InputMessageBody,
                ShowErrorAlertAfterInvalidDataIsEntered = ObjectForEquatableTests.ShowErrorAlertAfterInvalidDataIsEntered,
                ErrorAlertStyle = ObjectForEquatableTests.ErrorAlertStyle,
                ErrorAlertTitle = A.Dummy<string>(),
                ErrorAlertBody = ObjectForEquatableTests.ErrorAlertBody,
                ShowListDropdown = ObjectForEquatableTests.ShowListDropdown,
                Operand1NumericValue = ObjectForEquatableTests.Operand1NumericValue,
                Operand2NumericValue = ObjectForEquatableTests.Operand2NumericValue,
            },
            new NumericDataValidation
            {
                Kind = ObjectForEquatableTests.Kind,
                Operator = ObjectForEquatableTests.Operator,
                Operand1Formula = ObjectForEquatableTests.Operand1Formula,
                Operand2Formula = ObjectForEquatableTests.Operand2Formula,
                IgnoreBlank = ObjectForEquatableTests.IgnoreBlank,
                ShowInputMessage = ObjectForEquatableTests.ShowInputMessage,
                InputMessageTitle = ObjectForEquatableTests.InputMessageTitle,
                InputMessageBody = ObjectForEquatableTests.InputMessageBody,
                ShowErrorAlertAfterInvalidDataIsEntered = ObjectForEquatableTests.ShowErrorAlertAfterInvalidDataIsEntered,
                ErrorAlertStyle = ObjectForEquatableTests.ErrorAlertStyle,
                ErrorAlertTitle = ObjectForEquatableTests.ErrorAlertTitle,
                ErrorAlertBody = A.Dummy<string>(),
                ShowListDropdown = ObjectForEquatableTests.ShowListDropdown,
                Operand1NumericValue = ObjectForEquatableTests.Operand1NumericValue,
                Operand2NumericValue = ObjectForEquatableTests.Operand2NumericValue,
            },
            new NumericDataValidation
            {
                Kind = ObjectForEquatableTests.Kind,
                Operator = ObjectForEquatableTests.Operator,
                Operand1Formula = ObjectForEquatableTests.Operand1Formula,
                Operand2Formula = ObjectForEquatableTests.Operand2Formula,
                IgnoreBlank = ObjectForEquatableTests.IgnoreBlank,
                ShowInputMessage = ObjectForEquatableTests.ShowInputMessage,
                InputMessageTitle = ObjectForEquatableTests.InputMessageTitle,
                InputMessageBody = ObjectForEquatableTests.InputMessageBody,
                ShowErrorAlertAfterInvalidDataIsEntered = ObjectForEquatableTests.ShowErrorAlertAfterInvalidDataIsEntered,
                ErrorAlertStyle = ObjectForEquatableTests.ErrorAlertStyle,
                ErrorAlertTitle = ObjectForEquatableTests.ErrorAlertTitle,
                ErrorAlertBody = ObjectForEquatableTests.ErrorAlertBody,
                ShowListDropdown = !ObjectForEquatableTests.ShowListDropdown,
                Operand1NumericValue = ObjectForEquatableTests.Operand1NumericValue,
                Operand2NumericValue = ObjectForEquatableTests.Operand2NumericValue,
            },
            new NumericDataValidation
            {
                Kind = ObjectForEquatableTests.Kind,
                Operator = ObjectForEquatableTests.Operator,
                Operand1Formula = ObjectForEquatableTests.Operand1Formula,
                Operand2Formula = ObjectForEquatableTests.Operand2Formula,
                IgnoreBlank = ObjectForEquatableTests.IgnoreBlank,
                ShowInputMessage = ObjectForEquatableTests.ShowInputMessage,
                InputMessageTitle = ObjectForEquatableTests.InputMessageTitle,
                InputMessageBody = ObjectForEquatableTests.InputMessageBody,
                ShowErrorAlertAfterInvalidDataIsEntered = ObjectForEquatableTests.ShowErrorAlertAfterInvalidDataIsEntered,
                ErrorAlertStyle = ObjectForEquatableTests.ErrorAlertStyle,
                ErrorAlertTitle = ObjectForEquatableTests.ErrorAlertTitle,
                ErrorAlertBody = ObjectForEquatableTests.ErrorAlertBody,
                ShowListDropdown = ObjectForEquatableTests.ShowListDropdown,
                Operand1NumericValue = A.Dummy<long?>().ThatIsNot(ObjectForEquatableTests.Operand1NumericValue),
                Operand2NumericValue = ObjectForEquatableTests.Operand2NumericValue,
            },
            new NumericDataValidation
            {
                Kind = ObjectForEquatableTests.Kind,
                Operator = ObjectForEquatableTests.Operator,
                Operand1Formula = ObjectForEquatableTests.Operand1Formula,
                Operand2Formula = ObjectForEquatableTests.Operand2Formula,
                IgnoreBlank = ObjectForEquatableTests.IgnoreBlank,
                ShowInputMessage = ObjectForEquatableTests.ShowInputMessage,
                InputMessageTitle = ObjectForEquatableTests.InputMessageTitle,
                InputMessageBody = ObjectForEquatableTests.InputMessageBody,
                ShowErrorAlertAfterInvalidDataIsEntered = ObjectForEquatableTests.ShowErrorAlertAfterInvalidDataIsEntered,
                ErrorAlertStyle = ObjectForEquatableTests.ErrorAlertStyle,
                ErrorAlertTitle = ObjectForEquatableTests.ErrorAlertTitle,
                ErrorAlertBody = ObjectForEquatableTests.ErrorAlertBody,
                ShowListDropdown = ObjectForEquatableTests.ShowListDropdown,
                Operand1NumericValue = ObjectForEquatableTests.Operand1NumericValue,
                Operand2NumericValue = A.Dummy<long?>().ThatIsNot(ObjectForEquatableTests.Operand2NumericValue),
            },
        };

        private static readonly string ObjectThatIsNotTheSameTypeAsObjectForEquatableTests = A.Dummy<string>();

        [Fact]
        public static void Operand1Value___Should_be_same_as_Operand1NumericValue___When_getting()
        {
            // Arrange
            var systemUnderTest = A.Dummy<NumericDataValidation>();

            // Act
            var actual = systemUnderTest.Operand1Value;

            // Assert
            actual.Should().Be(systemUnderTest.Operand1NumericValue);
        }

        [Fact]
        public static void Operand2Value___Should_be_same_as_Operand2NumericValue___When_getting()
        {
            // Arrange
            var systemUnderTest = A.Dummy<NumericDataValidation>();

            // Act
            var actual = systemUnderTest.Operand2Value;

            // Assert
            actual.Should().Be(systemUnderTest.Operand2NumericValue);
        }

        [Fact]
        public static void Deserialize___Should_roundtrip_object___When_serializing_and_deserializing_using_ObcJsonSerializer()
        {
            // Arrange
            var expected = A.Dummy<NumericDataValidation>();
            var serializer = new ObcJsonSerializer(typeof(ExcelJsonConfiguration));
            var serializedJson = serializer.SerializeToString(expected);

            // Act
            var actual = serializer.Deserialize<NumericDataValidation>(serializedJson);

            // Assert
            actual.Should().Be(expected);
        }

        [Fact]
        public static void Deserialize___Should_roundtrip_object___When_serializing_and_deserializing_using_ObcBsonSerializer()
        {
            // Arrange
            var expected = A.Dummy<NumericDataValidation>();
            var serializer = new ObcBsonSerializer(configurationType: typeof(ExcelBsonConfiguration));

            var serializedBson = serializer.SerializeToString(expected);

            // Act
            var actual = serializer.Deserialize<NumericDataValidation>(serializedBson);

            // Assert
            actual.Should().Be(expected);
        }

        [Fact]
        public static void EqualsOperator___Should_return_true___When_both_sides_of_operator_are_null()
        {
            // Arrange
            NumericDataValidation systemUnderTest1 = null;
            NumericDataValidation systemUnderTest2 = null;

            // Act
            var result = systemUnderTest1 == systemUnderTest2;

            // Assert
            result.Should().BeTrue();
        }

        [Fact]
        public static void EqualsOperator___Should_return_false___When_one_side_of_operator_is_null_and_the_other_side_is_not_null()
        {
            // Arrange
            NumericDataValidation systemUnderTest = null;

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
            NumericDataValidation systemUnderTest1 = null;
            NumericDataValidation systemUnderTest2 = null;

            // Act
            var result = systemUnderTest1 != systemUnderTest2;

            // Assert
            result.Should().BeFalse();
        }

        [Fact]
        public static void NotEqualsOperator___Should_return_true___When_one_side_of_operator_is_null_and_the_other_side_is_not_null()
        {
            // Arrange
            NumericDataValidation systemUnderTest = null;

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
        public static void Equals_with_NumericDataValidation___Should_return_false___When_parameter_other_is_null()
        {
            // Arrange
            NumericDataValidation systemUnderTest = null;

            // Act
            var result = ObjectForEquatableTests.Equals(systemUnderTest);

            // Assert
            result.Should().BeFalse();
        }

        [Fact]
        public static void Equals_with_NumericDataValidation___Should_return_true___When_parameter_other_is_same_object()
        {
            // Arrange, Act
            var result = ObjectForEquatableTests.Equals(ObjectForEquatableTests);

            // Assert
            result.Should().BeTrue();
        }

        [Fact]
        public static void Equals_with_NumericDataValidation___Should_return_false___When_objects_being_compared_have_different_property_values()
        {
            // Arrange, Act
            var results = ObjectsThatAreNotEqualToObjectForEquatableTests.Select(_ => ObjectForEquatableTests.Equals(_)).ToList();

            // Assert
            results.ForEach(_ => _.Should().BeFalse());
        }

        [Fact]
        public static void Equals_with_NumericDataValidation___Should_return_true___When_objects_being_compared_have_same_property_values()
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
            var systemUnderTest1 = new NumericDataValidation();
            var systemUnderTest2 = A.Dummy<NumericDataValidation>();

            // Act
            var actual1 = systemUnderTest1.DeepClone();
            var actual2 = systemUnderTest2.DeepClone();

            // Assert
            actual1.Should().Be(systemUnderTest1);
            actual1.Should().NotBeSameAs(systemUnderTest1);

            actual2.Should().Be(systemUnderTest2);
            actual2.Should().NotBeSameAs(systemUnderTest2);
        }
    }
}
