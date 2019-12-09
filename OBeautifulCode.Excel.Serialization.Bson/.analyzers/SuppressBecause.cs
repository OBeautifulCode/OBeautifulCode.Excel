// --------------------------------------------------------------------------------------------------------------------
// <copyright file="SuppressBecause.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// <auto-generated>
//   Sourced from NuGet package. Will be overwritten with package update except in OBeautifulCode.Build source.
// </auto-generated>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.Serialization.Bson.Internal
{
    using System.CodeDom.Compiler;
    using System.Diagnostics.CodeAnalysis;
    
    /// <summary>
    /// Standard justifications for analysis suppression.
    /// </summary>
    [ExcludeFromCodeCoverage]
    [GeneratedCode("OBeautifulCode.Build.Analyzers", "See package version number")]
    internal static class ObcSuppressBecause
    {
        /// <summary>
        /// See the other suppression message(s) applied within the same context.
        /// </summary>
        public const string CA_ALL_SeeOtherSuppressionMessages = "See the other suppression messages applied within the same context.";

        /// <summary>
        /// The specified paramters are required to achieve the needed functionality.
        /// </summary>
        public const string CA1005_AvoidExcessiveParametersOnGenericTypes_SpecifiedParametersRequiredForNeededFunctionality = "The specified paramters are required to achieve the needed functionality.";

        /// <summary>
        /// Console executable does not need the [assembly: CLSCompliant(true)] as it should not be shared as an assembly for reference.
        /// </summary>
        public const string CA1014_MarkAssembliesWithClsCompliant_ConsoleExeDoesNotNeedToBeClsCompliant = "Console executable does not need the [assembly: CLSCompliant(true)] as it should not be shared as an assembly for reference.";

        /// <summary>
        /// We are optimizing for the logical grouping of types rather than the number of types in a namepace.
        /// </summary>
        public const string CA1020_AvoidNamespacesWithFewTypes_OptimizeForLogicalGroupingOfTypes = "We are optimizing for the logical grouping of types rather than the number of types in a namepace.";

        /// <summary>
        /// A visible nested type is required in unit tests.
        /// </summary>
        public const string CA1034_NestedTypesShouldNotBeVisible_VisibleNestedTypeRequiredForTesting = "A visible nested type is required in unit tests.";

        /// <summary>
        /// The type exists for unit tests that require a comparable type, but do not use the type to perform any comparisons.
        /// </summary>
        public const string CA1036_OverrideMethodsOnComparableTypes_TypeCreatedForTestsThatRequireComparableTypeButDoNotUseTypeToPerformComparisons = "The type exists for unit tests that require a comparable type, but do not use the type to perform any comparisons.";

        /// <summary>
        /// When we need to identify a group of types, we prefer the use of an empty interface over an attribute because it's easier to use and results in cleaner code.
        /// </summary>
        public const string CA1040_AvoidEmptyInterfaces_NeedToIdentifyGroupOfTypesAndPreferInterfaceOverAttribute = "When we need to identify a group of types, we prefer the use of an empty interface over an attribute because it's easier to use and results in cleaner code.";

        /// <summary>
        /// The type is used for test code that requires the instance field to be visible.
        /// </summary>
        public const string CA1051_DoNotDeclareVisibleInstanceFields_TypeUsedInTestingThatRequiresInstanceFieldToBeVisible = "The type is used for test code that requires the instance field to be visible.";

        /// <summary>
        /// It's ok to throw NotSupportedException for an unreachable code path.
        /// </summary>
        public const string CA1065_DoNotRaiseExceptionsInUnexpectedLocations_ThrowNotSupportedExceptionForUnreachableCodePath = "It's ok to throw NotSupportedException for an unreachable code path.";

        /// <summary>
        /// It's ok to throw NotImplementedException when a base type or implementing an interface forces us to create a member that will never be used in testing.
        /// </summary>
        public const string CA1065_DoNotRaiseExceptionsInUnexpectedLocations_ThrowNotImplementedExceptionWhenForcedToSpecifyMemberThatWillNeverBeUsedInTesting = "It's ok to throw NotImplementedException when a base type or implementing an interface forces us to create a member that will never be used in testing.";

        /// <summary>
        /// We prefer to read <see cref="System.Guid" />'s string representation as lowercase.
        /// </summary>
        public const string CA1308_NormalizeStringsToUppercase_PreferGuidLowercase = "We prefer to read System.Guid's string representation as lowercase.";

        /// <summary>
        /// We disagree with the assessment that this method as excessively complex.
        /// </summary>
        public const string CA1502_AvoidExcessiveComplexity_DisagreeWithAssessment = "We disagree with the assessment that this method as excessively complex.";

        /// <summary>
        /// The analyzer is incorrectly detecting compound words in a unit test method name.
        /// </summary>
        public const string CA1702_CompoundWordsShouldBeCasedCorrectly_AnalyzerIsIncorrectlyDetectingCompoundWordsInUnitTestMethodName = "The analyzer is incorrectly detecting compound words in a unit test method name.";

        /// <summary>
        /// The spelling of the identifier is correct in-context of the domain.
        /// </summary>
        public const string CA1704_IdentifiersShouldBeSpelledCorrectly_SpellingIsCorrectInContextOfTheDomain = "The spelling of the identifier is correct in-context of the domain.";

        /// <summary>
        /// The identifier is suffixed with the name of the Type that it directly extends or implements to improves readability and comprehension of unit tests whre the Type is a primary concern of those tests.
        /// </summary>
        public const string CA1710_IdentifiersShouldHaveCorrectSuffix_NameDirectlyExtendedOrImplementedTypeAddedAsSuffixForTestsWhereTypeIsPrimaryConcern = "The identifier is suffixed with the name of the Type that it directly extends or implements to improves readability and comprehension of unit tests whre the Type is a primary concern of those tests.";

        /// <summary>
        /// The identifier is suffixed with it's Type name to improve readability and comprehension of unit tests where the Type is a primary concern of those tests.
        /// </summary>
        public const string CA1711_IdentifiersShouldNotHaveIncorrectSuffix_TypeNameAddedAsSuffixForTestsWhereTypeIsPrimaryConcern = "The identifier is suffixed with it's Type name to improve readability and comprehension of unit tests where the Type is a primary concern of those tests.";

        /// <summary>
        /// The type name adds clarity to the identifier and there is no good alternative.
        /// </summary>
        public const string CA1720_IdentifiersShouldNotContainTypeNames_TypeNameAddsClarityToIdentifierAndNoGoodAlternative = "The type name adds clarity to the identifier and there is no good alternative.";

        /// <summary>
        /// The type name adds clarity to the identifier and the alternatives degrade the clarity of the identifier.
        /// </summary>
        public const string CA1720_IdentifiersShouldNotContainTypeNames_TypeNameAddsClarityToIdentifierAndAlternativesDegradeClarity = "The type name adds clarity to the identifier and the alternatives degrade the clarity of the identifier.";

        /// <summary>
        /// The identifier includes 'Flags' to improve readability and comprehension of unit tests where the kind of Enum is a primary concern of those tests.
        /// </summary>
        public const string CA1726_UsePreferredTerms_FlagsAddedForTestsWhereEnumKindIsPrimaryConcern = "The identifier includes 'Flags' to improve readability and comprehension of unit tests where the kind of Enum is a primary concern of those tests.";

        /// <summary>
        /// The type is being used in testing and we explicitly do not want the type to be equatable because it has bearing on the tests.
        /// </summary>
        public const string CA1815_OverrideEqualsAndOperatorEqualsOnValueTypes_TypeUsedForTestsThatRequireTypeToNotBeEquatable = "The type is being used in testing and we explicitly do not want the type to be equatable because it has bearing on the tests.";

        /// <summary>
        /// The type is immutable.
        /// </summary>
        public const string CA2104_DoNotDeclareReadOnlyMutableReferenceTypes_TypeIsImmutable = "The type is immutable.";

        /// <summary>
        /// The reserved exception is being used in unit test code; there is no real caller that will be impacted.
        /// </summary>
        public const string CA2201_DoNotRaiseReservedExceptionTypes_UsedForUnitTesting = "The reserved exception is being used in unit test code; there is no real caller that will be impacted.";

        /// <summary>
        /// The analyzer is incorectly flagging an object as being disposed multiple times.
        /// </summary>
        public const string CA2202_DoNotDisposeObjectsMultipleTimes_AnalyzerIsIncorrectlyFlaggingObjectAsBeingDisposedMultipleTimes = "The analyzer is incorectly flagging an object as being disposed multiple times.";

        /// <summary>
        /// The public interface of the system associated with this object never exposes this object.
        /// </summary>
        public const string CA2227_CollectionPropertiesShouldBeReadOnly_PublicInterfaceNeverExposesTheObject = "The public interface of the system associated with this object never exposes this object.";

        /// <summary>
        /// The type is used in unit tests with no intention to serialize.
        /// </summary>
        public const string CA2237_MarkISerializableTypesWithSerializable_UsedForTestingWithNoIntentionToSerialize = "The type is used in unit tests with no intention to serialize.";
    }
}
