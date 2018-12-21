// --------------------------------------------------------------------------------------------------------------------
// <copyright file="AsposeCellsLicense.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.AsposeCells
{
    using System;
    using System.IO;

    using Aspose.Cells;

    using OBeautifulCode.String.Recipes;
    using OBeautifulCode.Validation.Recipes;

    using static System.FormattableString;

    /// <summary>
    /// The Aspose.Cells license.
    /// </summary>
    public class AsposeCellsLicense
    {
        private static readonly object RegisterLock = new object();

        private static bool hasBeenRegistered = false;

        /// <summary>
        /// Initializes a new instance of the <see cref="AsposeCellsLicense"/> class.
        /// </summary>
        /// <param name="licenseXml">The license XML.</param>
        /// <exception cref="ArgumentNullException"><paramref name="licenseXml"/> is null.</exception>
        /// <exception cref="ArgumentException"><paramref name="licenseXml"/> is white space.</exception>
        public AsposeCellsLicense(
            string licenseXml)
        {
            new { licenseXml }.Must().NotBeNullNorWhiteSpace();

            this.LicenseXml = licenseXml;
        }

        /// <summary>
        /// Gets the license XML.
        /// </summary>
        public string LicenseXml { get; }

        /// <summary>
        /// Determines whether the Aspose.Cells license is registered.
        /// </summary>
        /// <returns>
        /// true if the Aspose.Cells license is registered; otherwise false.
        /// </returns>
        public static bool IsRegistered()
        {
            using (var workbook = new Workbook())
            {
                var result = workbook.IsLicensed;
                return result;
            }
        }

        /// <summary>
        /// Registers the license.
        /// </summary>
        /// <remarks>
        /// This method ensures that the license is only registered once per appdomain.
        /// </remarks>
        /// <exception cref="InvalidOperationException"><see cref="LicenseXml"/> is invalid or corrupt.</exception>
        public void Register()
        {
            lock (RegisterLock)
            {
                if (!hasBeenRegistered)
                {
                    var license = new License();

                    try
                    {
                        using (var ms = new MemoryStream(this.LicenseXml.ToUtf8Bytes()))
                        {
                            license.SetLicense(ms);
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException(Invariant($"{nameof(this.LicenseXml)} is invalid or corrupt.  See inner exception."), ex);
                    }

                    hasBeenRegistered = true;
                }
            }
        }
    }
}
