// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Packaging;
using System.IO;
using Xunit;

using static DocumentFormat.OpenXml.Tests.TestAssets;

namespace DocumentFormat.OpenXml.Tests
{
    public class CreateFromTemplateTests
    {
        [Fact]
        public void CanCreatePresentationFromTemplate()
        {
            using var stream = OpenFile(TestFiles.Templates.Presentation, FileAccess.ReadWrite);
            using var packageDocument = PresentationDocument.CreateFromTemplate(stream.Path);

            Assert.NotNull(packageDocument.PresentationPart);
            Assert.NotNull(packageDocument.PresentationPart!.Presentation);

            var clonePath = Path.GetTempFileName();
            try
            {
                using var clone = packageDocument.Clone(clonePath);
                Assert.NotNull(clone);
                Assert.True(new FileInfo(clonePath).Length > 0);
            }
            finally
            {
                File.Delete(clonePath);
            }
        }

        [Fact]
        public void CanCreateSpreadsheetFromTemplate()
        {
            using var stream = OpenFile(TestFiles.Templates.Spreadsheet, FileAccess.ReadWrite);
            using var packageDocument = SpreadsheetDocument.CreateFromTemplate(stream.Path);

            Assert.NotNull(packageDocument.WorkbookPart);
            Assert.NotNull(packageDocument.WorkbookPart!.Workbook);

            var clonePath = Path.GetTempFileName();
            try
            {
                using var clone = packageDocument.Clone(clonePath);
                Assert.NotNull(clone);
                Assert.True(new FileInfo(clonePath).Length > 0);
            }
            finally
            {
                File.Delete(clonePath);
            }
        }

        [Fact]
        public void CanCreateWordprocessingDocumentFromTemplate()
        {
            using var stream = OpenFile(TestFiles.Templates.Document, FileAccess.ReadWrite);
            using var packageDocument = WordprocessingDocument.CreateFromTemplate(stream.Path);

            Assert.NotNull(packageDocument.MainDocumentPart);
            Assert.NotNull(packageDocument.MainDocumentPart!.Document);

            var clonePath = Path.GetTempFileName();
            try
            {
                using var clone = packageDocument.Clone(clonePath);
                Assert.NotNull(clone);
                Assert.True(new FileInfo(clonePath).Length > 0);
            }
            finally
            {
                File.Delete(clonePath);
            }
        }
    }
}
