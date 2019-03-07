// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ExcelDummyFactory.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.Test
{
    using System;
    using System.Drawing;

    using FakeItEasy;

    using OBeautifulCode.AutoFakeItEasy;
    using OBeautifulCode.Math.Recipes;

    /// <inheritdoc />
    public class ExcelDummyFactory : IDummyFactory
    {
        public ExcelDummyFactory()
        {
            AutoFixtureBackedDummyFactory.UseRandomConcreteSubclassForDummy<DataValidation>();

            AutoFixtureBackedDummyFactory.ConstrainDummyToExclude(BorderEdges.Unknown);

            AutoFixtureBackedDummyFactory.AddDummyCreator(() =>
            {
                var result = Color.FromArgb(ThreadSafeRandom.Next(256), ThreadSafeRandom.Next(256), ThreadSafeRandom.Next(256));
                return result;
            });

            AutoFixtureBackedDummyFactory.AddDummyCreator(() =>
            {
                var worksheetName = "worksheet-" + A.Dummy<Guid>().ToString().Substring(1, 10);
                var rowNumber = A.Dummy<PositiveInteger>().ThatIs(_ => _ <= Constants.MaximumRowNumber);
                var columnNumber = A.Dummy<PositiveInteger>().ThatIs(_ => _ <= Constants.MaximumColumnNumber);
                var result = new CellReference(worksheetName, rowNumber, columnNumber);
                return result;
            });
        }

        /// <inheritdoc />
        public Priority Priority => new FakeItEasy.Priority(1);

        /// <inheritdoc />
        public bool CanCreate(Type type)
        {
            return false;
        }

        /// <inheritdoc />
        public object Create(Type type)
        {
            return null;
        }
    }
}