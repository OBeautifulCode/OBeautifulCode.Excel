﻿// --------------------------------------------------------------------------------------------------------------------
// <auto-generated>
//   Generated using OBeautifulCode.CodeGen.ModelObject (1.0.165.0)
// </auto-generated>
// --------------------------------------------------------------------------------------------------------------------

namespace OBeautifulCode.Excel.Test
{
    using global::System;
    using global::System.CodeDom.Compiler;
    using global::System.Collections.Concurrent;
    using global::System.Collections.Generic;
    using global::System.Collections.ObjectModel;
    using global::System.Diagnostics.CodeAnalysis;
    using global::System.Drawing;

    using global::FakeItEasy;

    using global::OBeautifulCode.AutoFakeItEasy;
    using global::OBeautifulCode.Excel;
    using global::OBeautifulCode.Math.Recipes;

    /// <summary>
    /// The default (code generated) Dummy Factory.
    /// Derive from this class to add any overriding or custom registrations.
    /// </summary>
    [ExcludeFromCodeCoverage]
    [GeneratedCode("OBeautifulCode.CodeGen.ModelObject", "1.0.165.0")]
#if !OBeautifulCodeExcelSolution
    internal
#else
    public
#endif
    abstract class DefaultExcelDummyFactory : IDummyFactory
    {
        public DefaultExcelDummyFactory()
        {
            AutoFixtureBackedDummyFactory.AddDummyCreator(
                () => new Border
                             {
                                 Edges = A.Dummy<BorderEdges>(),
                                 Style = A.Dummy<BorderStyle>(),
                                 Color = A.Dummy<Color>(),
                             });

            AutoFixtureBackedDummyFactory.AddDummyCreator(
                () => new CellReference(
                                 A.Dummy<string>(),
                                 A.Dummy<int>(),
                                 A.Dummy<int>()));

            AutoFixtureBackedDummyFactory.AddDummyCreator(
                () => new CellValueConditionalFormattingRule
                             {
                                 Operator        = A.Dummy<ConditionalFormattingOperator>(),
                                 Operand1Formula = A.Dummy<string>(),
                                 Operand2Formula = A.Dummy<string>(),
                                 StopIfTrue      = A.Dummy<bool>(),
                                 RangeStyle      = A.Dummy<RangeStyle>(),
                             });

            AutoFixtureBackedDummyFactory.AddDummyCreator(
                () => new Comment
                             {
                                 Body                 = A.Dummy<string>(),
                                 HtmlBody             = A.Dummy<string>(),
                                 FontName             = A.Dummy<string>(),
                                 FontColor            = A.Dummy<Color?>(),
                                 FontSize             = A.Dummy<int?>(),
                                 FontIsBold           = A.Dummy<bool?>(),
                                 HorizontalAlignment  = A.Dummy<HorizontalAlignment?>(),
                                 VerticalAlignment    = A.Dummy<VerticalAlignment?>(),
                                 AutoSize             = A.Dummy<bool?>(),
                                 HeightInInches       = A.Dummy<decimal?>(),
                                 WidthInInches        = A.Dummy<decimal?>(),
                                 FillColor            = A.Dummy<Color?>(),
                                 FillTransparency     = A.Dummy<decimal?>(),
                                 BorderColor          = A.Dummy<Color?>(),
                                 BorderStyle          = A.Dummy<CommentBorderStyle?>(),
                                 BorderWeightInPoints = A.Dummy<decimal?>(),
                             });

            AutoFixtureBackedDummyFactory.AddDummyCreator(
                () =>
                {
                    var availableTypes = new[]
                    {
                        typeof(NumericDataValidation),
                        typeof(TextDataValidation)
                    };

                    var randomIndex = ThreadSafeRandom.Next(0, availableTypes.Length);

                    var randomType = availableTypes[randomIndex];

                    var result = (DataValidation)AD.ummy(randomType);

                    return result;
                });

            AutoFixtureBackedDummyFactory.AddDummyCreator(
                () => new DocumentProperties
                             {
                                 BuiltInDocumentPropertyKindToValueMap = A.Dummy<IReadOnlyDictionary<BuiltInDocumentPropertyKind, string>>(),
                             });

            AutoFixtureBackedDummyFactory.AddDummyCreator(
                () => new NamedCell(
                                 A.Dummy<string>(),
                                 A.Dummy<CellReference>()));

            AutoFixtureBackedDummyFactory.AddDummyCreator(
                () => new NumericDataValidation
                             {
                                 Kind                                    = A.Dummy<DataValidationKind>(),
                                 Operator                                = A.Dummy<DataValidationOperator>(),
                                 Operand1Formula                         = A.Dummy<string>(),
                                 Operand2Formula                         = A.Dummy<string>(),
                                 IgnoreBlank                             = A.Dummy<bool>(),
                                 ShowInputMessage                        = A.Dummy<bool>(),
                                 InputMessageTitle                       = A.Dummy<string>(),
                                 InputMessageBody                        = A.Dummy<string>(),
                                 ShowErrorAlertAfterInvalidDataIsEntered = A.Dummy<bool>(),
                                 ErrorAlertStyle                         = A.Dummy<DataValidationErrorAlertStyle>(),
                                 ErrorAlertTitle                         = A.Dummy<string>(),
                                 ErrorAlertBody                          = A.Dummy<string>(),
                                 ShowListDropdown                        = A.Dummy<bool>(),
                                 Operand1NumericValue                    = A.Dummy<long?>(),
                                 Operand2NumericValue                    = A.Dummy<long?>(),
                             });

            AutoFixtureBackedDummyFactory.AddDummyCreator(
                () => new RangeStyle
                             {
                                 BackgroundColor     = A.Dummy<Color?>(),
                                 FontColor           = A.Dummy<Color?>(),
                                 FontName            = A.Dummy<string>(),
                                 FontSize            = A.Dummy<int?>(),
                                 FontIsItalic        = A.Dummy<bool?>(),
                                 FontIsBold          = A.Dummy<bool?>(),
                                 FontUnderline       = A.Dummy<UnderlineKind?>(),
                                 FontRotationAngle   = A.Dummy<int?>(),
                                 TextIsWrapped       = A.Dummy<bool?>(),
                                 IndentLevel         = A.Dummy<int?>(),
                                 RowHeightInPixels   = A.Dummy<int?>(),
                                 ColumnWidthInPixels = A.Dummy<int?>(),
                                 VerticalAlignment   = A.Dummy<VerticalAlignment?>(),
                                 HorizontalAlignment = A.Dummy<HorizontalAlignment?>(),
                                 MergeCells          = A.Dummy<bool?>(),
                                 AutofitRows         = A.Dummy<bool?>(),
                                 InsideBorder        = A.Dummy<Border>(),
                                 OutsideBorder       = A.Dummy<Border>(),
                                 DataValidation      = A.Dummy<DataValidation>(),
                                 Format              = A.Dummy<Format?>(),
                                 CustomFormatString  = A.Dummy<string>(),
                             });

            AutoFixtureBackedDummyFactory.AddDummyCreator(
                () => new TextDataValidation
                             {
                                 Kind                                    = A.Dummy<DataValidationKind>(),
                                 Operator                                = A.Dummy<DataValidationOperator>(),
                                 Operand1Formula                         = A.Dummy<string>(),
                                 Operand2Formula                         = A.Dummy<string>(),
                                 IgnoreBlank                             = A.Dummy<bool>(),
                                 ShowInputMessage                        = A.Dummy<bool>(),
                                 InputMessageTitle                       = A.Dummy<string>(),
                                 InputMessageBody                        = A.Dummy<string>(),
                                 ShowErrorAlertAfterInvalidDataIsEntered = A.Dummy<bool>(),
                                 ErrorAlertStyle                         = A.Dummy<DataValidationErrorAlertStyle>(),
                                 ErrorAlertTitle                         = A.Dummy<string>(),
                                 ErrorAlertBody                          = A.Dummy<string>(),
                                 ShowListDropdown                        = A.Dummy<bool>(),
                                 Operand1TextValue                       = A.Dummy<string>(),
                                 Operand2TextValue                       = A.Dummy<string>(),
                             });

            AutoFixtureBackedDummyFactory.AddDummyCreator(
                () => new WorkbookProtection
                             {
                                 ClearTextPassword = A.Dummy<string>(),
                             });

            AutoFixtureBackedDummyFactory.AddDummyCreator(
                () => new WorksheetProtection
                             {
                                 ClearTextPassword = A.Dummy<string>(),
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