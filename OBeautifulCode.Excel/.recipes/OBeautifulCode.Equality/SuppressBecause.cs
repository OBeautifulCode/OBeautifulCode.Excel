// --------------------------------------------------------------------------------------------------------------------
// <copyright file="SuppressBecause.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// <auto-generated>
//   Sourced from NuGet package. Will be overwritten with package update except in OBeautifulCode.Build source.
// </auto-generated>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Equality.Recipes.Internal
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
        /// Console executable does not need the [assembly: CLSCompliant(true)] as it should not be shared as an assembly for reference.
        /// </summary>
        public const string CA1014_MarkAssembliesWithClsCompliant_ConsoleExeDoesNotNeedToBeClsCompliant = "Console executable does not need the [assembly: CLSCompliant(true)] as it should not be shared as an assembly for reference.";

        /// <summary>
        /// We are optimizing for the logical grouping of types rather than the number of types in a namepace.
        /// </summary>
        public const string CA1020_AvoidNamespacesWithFewTypes_OptimizeForLogicalGroupingOfTypes = "We are optimizing for the logical grouping of types rather than the number of types in a namepace.";

        /// <summary>
        /// When we need to identify a group of types, we prefer the use of an empty interface over an attribute because it's easier to use and results in cleaner code.
        /// </summary>
        public const string CA1040_AvoidEmptyInterfaces_NeedToIdentifyGroupOfTypesAndPreferInterfaceOverAttribute = "When we need to identify a group of types, we prefer the use of an empty interface over an attribute because it's easier to use and results in cleaner code.";

        /// <summary>
        /// It's ok to throw NotSupportedException for an unreachable code path.
        /// </summary>
        public const string CA1065_DoNotRaiseExceptionsInUnexpectedLocations_ThrowNotSupportedExceptionForUnreachableCodePath = "It's ok to throw NotSupportedException for an unreachable code path.";

        /// <summary>
        /// We prefer to read <see cref="System.Guid" />'s string representation as lowercase.
        /// </summary>
        public const string CA1308_NormalizeStringsToUppercase_PreferGuidLowercase = "We prefer to read System.Guid's string representation as lowercase.";

        /// <summary>
        /// The analyzer is incorrectly detecting compound words in a unit test method name.
        /// </summary>
        public const string CA1702_CompoundWordsShouldBeCasedCorrectly_AnalyzerIsIncorrectlyDetectingCompoundWordsInUnitTestMethodName = "The analyzer is incorrectly detecting compound words in a unit test method name.";

        /// <summary>
        /// The spelling of the identifier is correct in-context of the domain.
        /// </summary>
        public const string CA1704_IdentifiersShouldBeSpelledCorrectly_SpellingIsCorrectInContextOfTheDomain = "The spelling of the identifier is correct in-context of the domain.";

        /// <summary>
        /// The type name adds clarity to the identifier and there is no good alternative.
        /// </summary>
        public const string CA1720_IdentifiersShouldNotContainTypeNames_TypeNameAddsClarityToIdentifyAndNoGoodAlternative = "The type name adds clarity to the identifier and there is no good alternative.";

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
    }
}
