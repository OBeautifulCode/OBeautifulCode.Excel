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
            if (licenseXml == null)
            {
                throw new ArgumentNullException(nameof(licenseXml));
            }

            if (string.IsNullOrWhiteSpace(licenseXml))
            {
                throw new ArgumentException(Invariant($"'{nameof(licenseXml)}' is white space"));
            }

            this.LicenseXml = licenseXml;
        }

        /// <summary>
        /// Gets the license XML.
        /// </summary>
        public string LicenseXml { get; }

        /// <summary>
        /// Determines whether an Aspose.Cells license is registered.
        /// </summary>
        /// <returns>
        /// true if an Aspose.Cells license is registered; otherwise false.
        /// </returns>
        public static bool IsRegistered()
        {
            lock (RegisterLock)
            {
                // not registered via this class?
                if (!hasBeenRegistered)
                {
                    // determine if registered in some other way
                    using (var workbook = new Workbook())
                    {
                        hasBeenRegistered = workbook.IsLicensed;
                    }
                }

                return hasBeenRegistered;
            }
        }

        /// <summary>
        /// Throws an <see cref="InvalidOperationException"/> if an Aspose.Cells
        /// license is not registered.
        /// </summary>
        public static void ThrowIfNotRegistered()
        {
            if (!IsRegistered())
            {
                throw new InvalidOperationException("An Aspose.Cells license has not been registered");
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
