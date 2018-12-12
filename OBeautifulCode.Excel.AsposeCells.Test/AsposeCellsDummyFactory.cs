// --------------------------------------------------------------------------------------------------------------------
// <copyright file="AsposeCellsDummyFactory.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.AsposeCells.Test
{
    using System;

    using Aspose.Cells;

    using FakeItEasy;

    using OBeautifulCode.AutoFakeItEasy;

    /// <inheritdoc />
    public class AsposeCellsDummyFactory : IDummyFactory
    {
        public AsposeCellsDummyFactory()
        {
            AutoFixtureBackedDummyFactory.AddDummyCreator(() =>
            {
                var workbook = new Workbook();

                var result = workbook.Worksheets[0];

                return result;
            });

            AutoFixtureBackedDummyFactory.AddDummyCreator(() =>
            {
                var result = new CellCursor(A.Dummy<Worksheet>(), A.Dummy<PositiveInteger>(), A.Dummy<PositiveInteger>());

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