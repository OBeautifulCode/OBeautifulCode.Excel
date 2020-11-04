﻿// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ReflectionHelper.Method.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// <auto-generated>
//   Sourced from NuGet package. Will be overwritten with package update except in OBeautifulCode.Reflection.Recipes source.
// </auto-generated>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Reflection.Recipes
{
    using global::System;
    using global::System.Collections.Generic;
    using global::System.Linq;
    using global::System.Reflection;

    using OBeautifulCode.Type.Recipes;

    using static global::System.FormattableString;

#if !OBeautifulCodeReflectionSolution
    internal
#else
    public
#endif
    static partial class ReflectionHelper
    {
        /// <summary>
        /// Gets the methods of the specified type,
        /// with various options to control the scope of methods included and optionally order the methods.
        /// </summary>
        /// <param name="type">The type.</param>
        /// <param name="memberRelationships">OPTIONAL value that scopes the search for members based on their relationship to <paramref name="type"/>.  DEFAULT is to include the members declared in or inherited by the specified type.</param>
        /// <param name="memberOwners">OPTIONAL value that scopes the search for members based on who owns the member.  DEFAULT is to include members owned by an object or owned by the type itself.</param>
        /// <param name="memberAccessModifiers">OPTIONAL value that scopes the search for members based on access modifiers.  DEFAULT is to include members having any supported access modifier.</param>
        /// <param name="memberAttributes">OPTIONAL value that scopes the search for members based on the presence or absence of certain attributes on those members.  DEFAULT is to include members that are not compiler generated.</param>
        /// <param name="orderMembersBy">OPTIONAL value that specifies how to the members.  DEFAULT is return the members in no particular order.</param>
        /// <returns>
        /// The methods in the specified order.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="type"/> is null.</exception>
        public static IReadOnlyList<MethodInfo> GetMethodsFiltered(
            this Type type,
            MemberRelationships memberRelationships = MemberRelationships.DeclaredOrInherited,
            MemberOwners memberOwners = MemberOwners.All,
            MemberAccessModifiers memberAccessModifiers = MemberAccessModifiers.All,
            MemberAttributes memberAttributes = MemberAttributes.NotCompilerGenerated,
            OrderMembersBy orderMembersBy = OrderMembersBy.None)
        {
            if (type == null)
            {
                throw new ArgumentNullException(nameof(type));
            }

            var result = type
                .GetMembersFiltered(memberRelationships, memberOwners, memberAccessModifiers, MemberKinds.Method, MemberMutability.All, memberAttributes, orderMembersBy)
                .Cast<MethodInfo>()
                .ToList();

            return result;
        }

        /// <summary>
        /// Gets the specified interface type's methods along with the methods of all implemented interfaces.
        /// </summary>
        /// <param name="interfaceType">The type of the interface.</param>
        /// <returns>
        /// The methods declared on the specified interface along with the methods of all implemented interfaces.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="interfaceType"/> is null.</exception>
        /// <exception cref="ArgumentException"><paramref name="interfaceType"/> is not an interface type.</exception>
        public static IReadOnlyCollection<MethodInfo> GetInterfaceDeclaredAndImplementedMethods(
            this Type interfaceType)
        {
            if (interfaceType == null)
            {
                throw new ArgumentNullException(nameof(interfaceType));
            }

            if (!interfaceType.IsInterface)
            {
                throw new ArgumentException(Invariant($"{nameof(interfaceType)} is not an interface type."));
            }

            var result = interfaceType.GetMethodsFiltered(MemberRelationships.DeclaredInTypeOrImplementedInterfaces).ToList();

            return result;
        }

        /// <summary>
        /// Gets the <see cref="MethodInfo"/> for the specified method.
        /// </summary>
        /// <param name="type">The type.</param>
        /// <param name="methodName">The name of the method.</param>
        /// <param name="memberRelationships">OPTIONAL value that scopes the search for members based on their relationship to <paramref name="type"/>.  DEFAULT is to include the members declared in or inherited by the specified type.</param>
        /// <param name="memberOwners">OPTIONAL value that scopes the search for members based on who owns the member.  DEFAULT is to include members owned by an object or owned by the type itself.</param>
        /// <param name="memberAccessModifiers">OPTIONAL value that scopes the search for members based on access modifiers.  DEFAULT is to include members having any supported access modifier.</param>
        /// <param name="memberAttributes">OPTIONAL value that scopes the search for members based on the presence or absence of certain attributes on those members.  DEFAULT is to include members that are not compiler generated.</param>
        /// <param name="throwIfNotFound">OPTIONAL value indicating whether to throw if no methods are found.  DEFAULT is to throw..</param>
        /// <returns>
        /// The <see cref="MethodInfo"/> or null if no methods are found and <paramref name="throwIfNotFound"/> is false
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="type"/> is null.</exception>
        /// <exception cref="ArgumentNullException"><paramref name="methodName"/> is null.</exception>
        /// <exception cref="ArgumentException"><paramref name="methodName"/> is whitespace.</exception>
        /// <exception cref="ArgumentException">There is no method named <paramref name="methodName"/> on the object type using the specified binding constraints and <paramref name="throwIfNotFound"/> is true.</exception>
        /// <exception cref="ArgumentException">There is more than one method named <paramref name="methodName"/> on the object type using the specified binding constraints.</exception>
        public static MethodInfo GetMethodFiltered(
            this Type type,
            string methodName,
            MemberRelationships memberRelationships = MemberRelationships.DeclaredOrInherited,
            MemberOwners memberOwners = MemberOwners.All,
            MemberAccessModifiers memberAccessModifiers = MemberAccessModifiers.All,
            MemberAttributes memberAttributes = MemberAttributes.NotCompilerGenerated,
            bool throwIfNotFound = true)
        {
            if (type == null)
            {
                throw new ArgumentNullException(nameof(type));
            }

            if (methodName == null)
            {
                throw new ArgumentNullException(nameof(methodName));
            }

            if (string.IsNullOrWhiteSpace(methodName))
            {
                throw new ArgumentException(Invariant($"{nameof(methodName)} is white space."));
            }

            var methods = type
                // ReSharper disable once RedundantArgumentDefaultValue
                .GetMethodsFiltered(memberRelationships, memberOwners, memberAccessModifiers, memberAttributes, OrderMembersBy.None)
                .Where(_ => _.Name == methodName)
                .ToList();

            MethodInfo result;

            if (!methods.Any())
            {
                if (throwIfNotFound)
                {
                    throw new ArgumentException(Invariant($"There is no method named '{methodName}' on type '{type.ToStringReadable()}', using the specified binding constraints."));
                }
                else
                {
                    result = null;
                }
            }
            else if (methods.Count > 1)
            {
                throw new ArgumentException(Invariant($"There is more than one method named '{methodName}' on type '{type.ToStringReadable()}', using the specified binding constraints."));
            }
            else
            {
                result = methods.Single();
            }

            return result;
        }

        /// <summary>
        /// Determines if a type has a method of the specified method name.
        /// </summary>
        /// <param name="type">The type to check.</param>
        /// <param name="methodName">The name of the method to check for.</param>
        /// <param name="memberRelationships">OPTIONAL value that scopes the search for members based on their relationship to <paramref name="type"/>.  DEFAULT is to include the members declared in or inherited by the specified type.</param>
        /// <param name="memberOwners">OPTIONAL value that scopes the search for members based on who owns the member.  DEFAULT is to include members owned by an object or owned by the type itself.</param>
        /// <param name="memberAccessModifiers">OPTIONAL value that scopes the search for members based on access modifiers.  DEFAULT is to include members having any supported access modifier.</param>
        /// <param name="memberAttributes">OPTIONAL value that scopes the search for members based on the presence or absence of certain attributes on those members.  DEFAULT is to include members that are not compiler generated.</param>
        /// <returns>
        /// true if the type has a method of the specified method name, false if not.
        /// </returns>
        /// <exception cref="ArgumentNullException"><paramref name="type"/> is null.</exception>
        /// <exception cref="ArgumentNullException"><paramref name="methodName"/> is null.</exception>
        /// <exception cref="ArgumentException"><paramref name="methodName"/> is whitespace.</exception>
        public static bool HasMethod(
            this Type type,
            string methodName,
            MemberRelationships memberRelationships = MemberRelationships.DeclaredOrInherited,
            MemberOwners memberOwners = MemberOwners.All,
            MemberAccessModifiers memberAccessModifiers = MemberAccessModifiers.All,
            MemberAttributes memberAttributes = MemberAttributes.NotCompilerGenerated)
        {
            if (type == null)
            {
                throw new ArgumentNullException(nameof(type));
            }

            if (methodName == null)
            {
                throw new ArgumentNullException(nameof(methodName));
            }

            if (string.IsNullOrWhiteSpace(methodName))
            {
                throw new ArgumentException(Invariant($"{nameof(methodName)} is white space."));
            }

            var methods = type
                // ReSharper disable once RedundantArgumentDefaultValue
                .GetMethodsFiltered(memberRelationships, memberOwners, memberAccessModifiers, memberAttributes, OrderMembersBy.None)
                .Where(_ => _.Name == methodName)
                .ToList();

            var result = methods.Any();

            return result;
        }
    }
}
