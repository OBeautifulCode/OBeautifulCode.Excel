// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CommentTest.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.Test.Style
{
    using System.Drawing;
    using System.Linq;

    using FakeItEasy;

    using FluentAssertions;

    using Naos.Serialization.Bson;
    using Naos.Serialization.Json;

    using OBeautifulCode.AutoFakeItEasy;
    using OBeautifulCode.Excel.Serialization.Bson;

    using Xunit;

    public static class CommentTest
    {
        private static readonly Comment ObjectForEquatableTests = A.Dummy<Comment>();

        private static readonly Comment ObjectThatIsEqualButNotTheSameAsObjectForEquatableTests =
            new Comment
            {
                Body = ObjectForEquatableTests.Body,
                HtmlBody = ObjectForEquatableTests.HtmlBody,
                FontName = ObjectForEquatableTests.FontName,
                FontColor = ObjectForEquatableTests.FontColor,
                FontSize = ObjectForEquatableTests.FontSize,
                FontIsBold = ObjectForEquatableTests.FontIsBold,
                HorizontalAlignment = ObjectForEquatableTests.HorizontalAlignment,
                VerticalAlignment = ObjectForEquatableTests.VerticalAlignment,
                AutoSize = ObjectForEquatableTests.AutoSize,
                HeightInInches = ObjectForEquatableTests.HeightInInches,
                WidthInInches = ObjectForEquatableTests.WidthInInches,
                FillColor = ObjectForEquatableTests.FillColor,
                FillTransparency = ObjectForEquatableTests.FillTransparency,
                BorderColor = ObjectForEquatableTests.BorderColor,
                BorderStyle = ObjectForEquatableTests.BorderStyle,
                BorderWeightInPoints = ObjectForEquatableTests.BorderWeightInPoints,
            };

        private static readonly Comment[] ObjectsThatAreNotEqualToObjectForEquatableTests =
        {
            A.Dummy<Comment>(),
            new Comment
            {
                Body = A.Dummy<string>(),
                HtmlBody = ObjectForEquatableTests.HtmlBody,
                FontName = ObjectForEquatableTests.FontName,
                FontColor = ObjectForEquatableTests.FontColor,
                FontSize = ObjectForEquatableTests.FontSize,
                FontIsBold = ObjectForEquatableTests.FontIsBold,
                HorizontalAlignment = ObjectForEquatableTests.HorizontalAlignment,
                VerticalAlignment = ObjectForEquatableTests.VerticalAlignment,
                AutoSize = ObjectForEquatableTests.AutoSize,
                HeightInInches = ObjectForEquatableTests.HeightInInches,
                WidthInInches = ObjectForEquatableTests.WidthInInches,
                FillColor = ObjectForEquatableTests.FillColor,
                FillTransparency = ObjectForEquatableTests.FillTransparency,
                BorderColor = ObjectForEquatableTests.BorderColor,
                BorderStyle = ObjectForEquatableTests.BorderStyle,
                BorderWeightInPoints = ObjectForEquatableTests.BorderWeightInPoints,
            },
            new Comment
            {
                Body = ObjectForEquatableTests.Body,
                HtmlBody = A.Dummy<string>(),
                FontName = ObjectForEquatableTests.FontName,
                FontColor = ObjectForEquatableTests.FontColor,
                FontSize = ObjectForEquatableTests.FontSize,
                FontIsBold = ObjectForEquatableTests.FontIsBold,
                HorizontalAlignment = ObjectForEquatableTests.HorizontalAlignment,
                VerticalAlignment = ObjectForEquatableTests.VerticalAlignment,
                AutoSize = ObjectForEquatableTests.AutoSize,
                HeightInInches = ObjectForEquatableTests.HeightInInches,
                WidthInInches = ObjectForEquatableTests.WidthInInches,
                FillColor = ObjectForEquatableTests.FillColor,
                FillTransparency = ObjectForEquatableTests.FillTransparency,
                BorderColor = ObjectForEquatableTests.BorderColor,
                BorderStyle = ObjectForEquatableTests.BorderStyle,
                BorderWeightInPoints = ObjectForEquatableTests.BorderWeightInPoints,
            },
            new Comment
            {
                Body = ObjectForEquatableTests.Body,
                HtmlBody = ObjectForEquatableTests.HtmlBody,
                FontName = A.Dummy<string>(),
                FontColor = ObjectForEquatableTests.FontColor,
                FontSize = ObjectForEquatableTests.FontSize,
                FontIsBold = ObjectForEquatableTests.FontIsBold,
                HorizontalAlignment = ObjectForEquatableTests.HorizontalAlignment,
                VerticalAlignment = ObjectForEquatableTests.VerticalAlignment,
                AutoSize = ObjectForEquatableTests.AutoSize,
                HeightInInches = ObjectForEquatableTests.HeightInInches,
                WidthInInches = ObjectForEquatableTests.WidthInInches,
                FillColor = ObjectForEquatableTests.FillColor,
                FillTransparency = ObjectForEquatableTests.FillTransparency,
                BorderColor = ObjectForEquatableTests.BorderColor,
                BorderStyle = ObjectForEquatableTests.BorderStyle,
                BorderWeightInPoints = ObjectForEquatableTests.BorderWeightInPoints,
            },
            new Comment
            {
                Body = ObjectForEquatableTests.Body,
                HtmlBody = ObjectForEquatableTests.HtmlBody,
                FontName = ObjectForEquatableTests.FontName,
                FontColor = A.Dummy<Color>().ThatIs(_ => _ != ObjectForEquatableTests.FontColor),
                FontSize = ObjectForEquatableTests.FontSize,
                FontIsBold = ObjectForEquatableTests.FontIsBold,
                HorizontalAlignment = ObjectForEquatableTests.HorizontalAlignment,
                VerticalAlignment = ObjectForEquatableTests.VerticalAlignment,
                AutoSize = ObjectForEquatableTests.AutoSize,
                HeightInInches = ObjectForEquatableTests.HeightInInches,
                WidthInInches = ObjectForEquatableTests.WidthInInches,
                FillColor = ObjectForEquatableTests.FillColor,
                FillTransparency = ObjectForEquatableTests.FillTransparency,
                BorderColor = ObjectForEquatableTests.BorderColor,
                BorderStyle = ObjectForEquatableTests.BorderStyle,
                BorderWeightInPoints = ObjectForEquatableTests.BorderWeightInPoints,
            },
            new Comment
            {
                Body = ObjectForEquatableTests.Body,
                HtmlBody = ObjectForEquatableTests.HtmlBody,
                FontName = ObjectForEquatableTests.FontName,
                FontColor = ObjectForEquatableTests.FontColor,
                FontSize = A.Dummy<int>(),
                FontIsBold = ObjectForEquatableTests.FontIsBold,
                HorizontalAlignment = ObjectForEquatableTests.HorizontalAlignment,
                VerticalAlignment = ObjectForEquatableTests.VerticalAlignment,
                AutoSize = ObjectForEquatableTests.AutoSize,
                HeightInInches = ObjectForEquatableTests.HeightInInches,
                WidthInInches = ObjectForEquatableTests.WidthInInches,
                FillColor = ObjectForEquatableTests.FillColor,
                FillTransparency = ObjectForEquatableTests.FillTransparency,
                BorderColor = ObjectForEquatableTests.BorderColor,
                BorderStyle = ObjectForEquatableTests.BorderStyle,
                BorderWeightInPoints = ObjectForEquatableTests.BorderWeightInPoints,
            },
            new Comment
            {
                Body = ObjectForEquatableTests.Body,
                HtmlBody = ObjectForEquatableTests.HtmlBody,
                FontName = ObjectForEquatableTests.FontName,
                FontColor = ObjectForEquatableTests.FontColor,
                FontSize = ObjectForEquatableTests.FontSize,
                FontIsBold = !ObjectForEquatableTests.FontIsBold,
                HorizontalAlignment = ObjectForEquatableTests.HorizontalAlignment,
                VerticalAlignment = ObjectForEquatableTests.VerticalAlignment,
                AutoSize = ObjectForEquatableTests.AutoSize,
                HeightInInches = ObjectForEquatableTests.HeightInInches,
                WidthInInches = ObjectForEquatableTests.WidthInInches,
                FillColor = ObjectForEquatableTests.FillColor,
                FillTransparency = ObjectForEquatableTests.FillTransparency,
                BorderColor = ObjectForEquatableTests.BorderColor,
                BorderStyle = ObjectForEquatableTests.BorderStyle,
                BorderWeightInPoints = ObjectForEquatableTests.BorderWeightInPoints,
            },
            new Comment
            {
                Body = ObjectForEquatableTests.Body,
                HtmlBody = ObjectForEquatableTests.HtmlBody,
                FontName = ObjectForEquatableTests.FontName,
                FontColor = ObjectForEquatableTests.FontColor,
                FontSize = ObjectForEquatableTests.FontSize,
                FontIsBold = ObjectForEquatableTests.FontIsBold,
                HorizontalAlignment = A.Dummy<HorizontalAlignment?>().ThatIsNot(ObjectForEquatableTests.HorizontalAlignment),
                VerticalAlignment = ObjectForEquatableTests.VerticalAlignment,
                AutoSize = ObjectForEquatableTests.AutoSize,
                HeightInInches = ObjectForEquatableTests.HeightInInches,
                WidthInInches = ObjectForEquatableTests.WidthInInches,
                FillColor = ObjectForEquatableTests.FillColor,
                FillTransparency = ObjectForEquatableTests.FillTransparency,
                BorderColor = ObjectForEquatableTests.BorderColor,
                BorderStyle = ObjectForEquatableTests.BorderStyle,
                BorderWeightInPoints = ObjectForEquatableTests.BorderWeightInPoints,
            },
            new Comment
            {
                Body = ObjectForEquatableTests.Body,
                HtmlBody = ObjectForEquatableTests.HtmlBody,
                FontName = ObjectForEquatableTests.FontName,
                FontColor = ObjectForEquatableTests.FontColor,
                FontSize = ObjectForEquatableTests.FontSize,
                FontIsBold = ObjectForEquatableTests.FontIsBold,
                HorizontalAlignment = ObjectForEquatableTests.HorizontalAlignment,
                VerticalAlignment = A.Dummy<VerticalAlignment?>().ThatIsNot(ObjectForEquatableTests.VerticalAlignment),
                AutoSize = ObjectForEquatableTests.AutoSize,
                HeightInInches = ObjectForEquatableTests.HeightInInches,
                WidthInInches = ObjectForEquatableTests.WidthInInches,
                FillColor = ObjectForEquatableTests.FillColor,
                FillTransparency = ObjectForEquatableTests.FillTransparency,
                BorderColor = ObjectForEquatableTests.BorderColor,
                BorderStyle = ObjectForEquatableTests.BorderStyle,
                BorderWeightInPoints = ObjectForEquatableTests.BorderWeightInPoints,
            },
            new Comment
            {
                Body = ObjectForEquatableTests.Body,
                HtmlBody = ObjectForEquatableTests.HtmlBody,
                FontName = ObjectForEquatableTests.FontName,
                FontColor = ObjectForEquatableTests.FontColor,
                FontSize = ObjectForEquatableTests.FontSize,
                FontIsBold = ObjectForEquatableTests.FontIsBold,
                HorizontalAlignment = ObjectForEquatableTests.HorizontalAlignment,
                VerticalAlignment = ObjectForEquatableTests.VerticalAlignment,
                AutoSize = !ObjectForEquatableTests.AutoSize,
                HeightInInches = ObjectForEquatableTests.HeightInInches,
                WidthInInches = ObjectForEquatableTests.WidthInInches,
                FillColor = ObjectForEquatableTests.FillColor,
                FillTransparency = ObjectForEquatableTests.FillTransparency,
                BorderColor = ObjectForEquatableTests.BorderColor,
                BorderStyle = ObjectForEquatableTests.BorderStyle,
                BorderWeightInPoints = ObjectForEquatableTests.BorderWeightInPoints,
            },
            new Comment
            {
                Body = ObjectForEquatableTests.Body,
                HtmlBody = ObjectForEquatableTests.HtmlBody,
                FontName = ObjectForEquatableTests.FontName,
                FontColor = ObjectForEquatableTests.FontColor,
                FontSize = ObjectForEquatableTests.FontSize,
                FontIsBold = ObjectForEquatableTests.FontIsBold,
                HorizontalAlignment = ObjectForEquatableTests.HorizontalAlignment,
                VerticalAlignment = ObjectForEquatableTests.VerticalAlignment,
                AutoSize = ObjectForEquatableTests.AutoSize,
                HeightInInches = A.Dummy<decimal>(),
                WidthInInches = ObjectForEquatableTests.WidthInInches,
                FillColor = ObjectForEquatableTests.FillColor,
                FillTransparency = ObjectForEquatableTests.FillTransparency,
                BorderColor = ObjectForEquatableTests.BorderColor,
                BorderStyle = ObjectForEquatableTests.BorderStyle,
                BorderWeightInPoints = ObjectForEquatableTests.BorderWeightInPoints,
            },
            new Comment
            {
                Body = ObjectForEquatableTests.Body,
                HtmlBody = ObjectForEquatableTests.HtmlBody,
                FontName = ObjectForEquatableTests.FontName,
                FontColor = ObjectForEquatableTests.FontColor,
                FontSize = ObjectForEquatableTests.FontSize,
                FontIsBold = ObjectForEquatableTests.FontIsBold,
                HorizontalAlignment = ObjectForEquatableTests.HorizontalAlignment,
                VerticalAlignment = ObjectForEquatableTests.VerticalAlignment,
                AutoSize = ObjectForEquatableTests.AutoSize,
                HeightInInches = ObjectForEquatableTests.HeightInInches,
                WidthInInches = A.Dummy<decimal>(),
                FillColor = ObjectForEquatableTests.FillColor,
                FillTransparency = ObjectForEquatableTests.FillTransparency,
                BorderColor = ObjectForEquatableTests.BorderColor,
                BorderStyle = ObjectForEquatableTests.BorderStyle,
                BorderWeightInPoints = ObjectForEquatableTests.BorderWeightInPoints,
            },
            new Comment
            {
                Body = ObjectForEquatableTests.Body,
                HtmlBody = ObjectForEquatableTests.HtmlBody,
                FontName = ObjectForEquatableTests.FontName,
                FontColor = ObjectForEquatableTests.FontColor,
                FontSize = ObjectForEquatableTests.FontSize,
                FontIsBold = ObjectForEquatableTests.FontIsBold,
                HorizontalAlignment = ObjectForEquatableTests.HorizontalAlignment,
                VerticalAlignment = ObjectForEquatableTests.VerticalAlignment,
                AutoSize = ObjectForEquatableTests.AutoSize,
                HeightInInches = ObjectForEquatableTests.HeightInInches,
                WidthInInches = ObjectForEquatableTests.WidthInInches,
                FillColor = A.Dummy<Color?>().ThatIsNot(ObjectForEquatableTests.FillColor),
                FillTransparency = ObjectForEquatableTests.FillTransparency,
                BorderColor = ObjectForEquatableTests.BorderColor,
                BorderStyle = ObjectForEquatableTests.BorderStyle,
                BorderWeightInPoints = ObjectForEquatableTests.BorderWeightInPoints,
            },
            new Comment
            {
                Body = ObjectForEquatableTests.Body,
                HtmlBody = ObjectForEquatableTests.HtmlBody,
                FontName = ObjectForEquatableTests.FontName,
                FontColor = ObjectForEquatableTests.FontColor,
                FontSize = ObjectForEquatableTests.FontSize,
                FontIsBold = ObjectForEquatableTests.FontIsBold,
                HorizontalAlignment = ObjectForEquatableTests.HorizontalAlignment,
                VerticalAlignment = ObjectForEquatableTests.VerticalAlignment,
                AutoSize = ObjectForEquatableTests.AutoSize,
                HeightInInches = ObjectForEquatableTests.HeightInInches,
                WidthInInches = ObjectForEquatableTests.WidthInInches,
                FillColor = ObjectForEquatableTests.FillColor,
                FillTransparency = A.Dummy<decimal>(),
                BorderColor = ObjectForEquatableTests.BorderColor,
                BorderStyle = ObjectForEquatableTests.BorderStyle,
                BorderWeightInPoints = ObjectForEquatableTests.BorderWeightInPoints,
            },
            new Comment
            {
                Body = ObjectForEquatableTests.Body,
                HtmlBody = ObjectForEquatableTests.HtmlBody,
                FontName = ObjectForEquatableTests.FontName,
                FontColor = ObjectForEquatableTests.FontColor,
                FontSize = ObjectForEquatableTests.FontSize,
                FontIsBold = ObjectForEquatableTests.FontIsBold,
                HorizontalAlignment = ObjectForEquatableTests.HorizontalAlignment,
                VerticalAlignment = ObjectForEquatableTests.VerticalAlignment,
                AutoSize = ObjectForEquatableTests.AutoSize,
                HeightInInches = ObjectForEquatableTests.HeightInInches,
                WidthInInches = ObjectForEquatableTests.WidthInInches,
                FillColor = ObjectForEquatableTests.FillColor,
                FillTransparency = ObjectForEquatableTests.FillTransparency,
                BorderColor = A.Dummy<Color?>().ThatIsNot(ObjectForEquatableTests.BorderColor),
                BorderStyle = ObjectForEquatableTests.BorderStyle,
                BorderWeightInPoints = ObjectForEquatableTests.BorderWeightInPoints,
            },
            new Comment
            {
                Body = ObjectForEquatableTests.Body,
                HtmlBody = ObjectForEquatableTests.HtmlBody,
                FontName = ObjectForEquatableTests.FontName,
                FontColor = ObjectForEquatableTests.FontColor,
                FontSize = ObjectForEquatableTests.FontSize,
                FontIsBold = ObjectForEquatableTests.FontIsBold,
                HorizontalAlignment = ObjectForEquatableTests.HorizontalAlignment,
                VerticalAlignment = ObjectForEquatableTests.VerticalAlignment,
                AutoSize = ObjectForEquatableTests.AutoSize,
                HeightInInches = ObjectForEquatableTests.HeightInInches,
                WidthInInches = ObjectForEquatableTests.WidthInInches,
                FillColor = ObjectForEquatableTests.FillColor,
                FillTransparency = ObjectForEquatableTests.FillTransparency,
                BorderColor = ObjectForEquatableTests.BorderColor,
                BorderStyle = A.Dummy<CommentBorderStyle?>().ThatIsNot(ObjectForEquatableTests.BorderStyle),
                BorderWeightInPoints = ObjectForEquatableTests.BorderWeightInPoints,
            },
            new Comment
            {
                Body = ObjectForEquatableTests.Body,
                HtmlBody = ObjectForEquatableTests.HtmlBody,
                FontName = ObjectForEquatableTests.FontName,
                FontColor = ObjectForEquatableTests.FontColor,
                FontSize = ObjectForEquatableTests.FontSize,
                FontIsBold = ObjectForEquatableTests.FontIsBold,
                HorizontalAlignment = ObjectForEquatableTests.HorizontalAlignment,
                VerticalAlignment = ObjectForEquatableTests.VerticalAlignment,
                AutoSize = ObjectForEquatableTests.AutoSize,
                HeightInInches = ObjectForEquatableTests.HeightInInches,
                WidthInInches = ObjectForEquatableTests.WidthInInches,
                FillColor = ObjectForEquatableTests.FillColor,
                FillTransparency = ObjectForEquatableTests.FillTransparency,
                BorderColor = ObjectForEquatableTests.BorderColor,
                BorderStyle = ObjectForEquatableTests.BorderStyle,
                BorderWeightInPoints = A.Dummy<decimal>(),
            },
        };

        private static readonly string ObjectThatIsNotTheSameTypeAsObjectForEquatableTests = A.Dummy<string>();

        [Fact]
        public static void Deserialize___Should_roundtrip_object___When_serializing_and_deserializing_using_NaosJsonSerializer()
        {
            // Arrange
            var expected = A.Dummy<Comment>();
            var serializer = new NaosJsonSerializer();
            var serializedJson = serializer.SerializeToString(expected);

            // Act
            var actual = serializer.Deserialize<Comment>(serializedJson);

            // Assert
            actual.Should().Be(expected);
        }

        [Fact]
        public static void Deserialize___Should_roundtrip_object___When_serializing_and_deserializing_using_NaosBsonSerializer()
        {
            // Arrange
            var expected = A.Dummy<Comment>();
            var serializer = new NaosBsonSerializer(configurationType: typeof(ExcelBsonConfiguration));

            var serializedJson = serializer.SerializeToString(expected);

            // Act
            var actual = serializer.Deserialize<Comment>(serializedJson);

            // Assert
            actual.Should().Be(expected);
        }

        [Fact]
        public static void EqualsOperator___Should_return_true___When_both_sides_of_operator_are_null()
        {
            // Arrange
            Comment systemUnderTest1 = null;
            Comment systemUnderTest2 = null;

            // Act
            var result = systemUnderTest1 == systemUnderTest2;

            // Assert
            result.Should().BeTrue();
        }

        [Fact]
        public static void EqualsOperator___Should_return_false___When_one_side_of_operator_is_null_and_the_other_side_is_not_null()
        {
            // Arrange
            Comment systemUnderTest = null;

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
            Comment systemUnderTest1 = null;
            Comment systemUnderTest2 = null;

            // Act
            var result = systemUnderTest1 != systemUnderTest2;

            // Assert
            result.Should().BeFalse();
        }

        [Fact]
        public static void NotEqualsOperator___Should_return_true___When_one_side_of_operator_is_null_and_the_other_side_is_not_null()
        {
            // Arrange
            Comment systemUnderTest = null;

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
        public static void Equals_with_Comment___Should_return_false___When_parameter_other_is_null()
        {
            // Arrange
            Comment systemUnderTest = null;

            // Act
            var result = ObjectForEquatableTests.Equals(systemUnderTest);

            // Assert
            result.Should().BeFalse();
        }

        [Fact]
        public static void Equals_with_Comment___Should_return_true___When_parameter_other_is_same_object()
        {
            // Arrange, Act
            var result = ObjectForEquatableTests.Equals(ObjectForEquatableTests);

            // Assert
            result.Should().BeTrue();
        }

        [Fact]
        public static void Equals_with_Comment___Should_return_false___When_objects_being_compared_have_different_property_values()
        {
            // Arrange, Act
            var results = ObjectsThatAreNotEqualToObjectForEquatableTests.Select(_ => ObjectForEquatableTests.Equals(_)).ToList();

            // Assert
            results.ForEach(_ => _.Should().BeFalse());
        }

        [Fact]
        public static void Equals_with_Comment___Should_return_true___When_objects_being_compared_have_same_property_values()
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
        public static void Clone___Should_clone_item___When_called()
        {
            // Arrange
            var systemUnderTest = A.Dummy<Comment>();

            // Act
            var actual = systemUnderTest.Clone();

            // Assert
            actual.Should().Be(systemUnderTest);
            actual.Should().NotBeSameAs(systemUnderTest);
        }
    }
}
