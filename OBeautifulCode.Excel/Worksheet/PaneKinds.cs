// --------------------------------------------------------------------------------------------------------------------
// <copyright file="PaneKinds.cs" company="OBeautifulCode">
//   Copyright (c) OBeautifulCode 2018. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel
{
    using System;

    /// <summary>
    /// Determines the kinds of pane.
    /// </summary>
    [Flags]
    public enum PaneKinds
    {
        /// <summary>
        /// No panes.
        /// </summary>
        None = 0,

        /// <summary>
        /// Row pane.
        /// </summary>
        Row = 1,

        /// <summary>
        /// Column pane.
        /// </summary>
        Column = 2,
    }
}
