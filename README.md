# Word Document Generation Using Syncfusion DocIO (VB.NET)

A pure .NET console application written in **VB.NET** demonstrating how to programmatically create and format Microsoft Word documents (`.docx`) using **[Syncfusion DocIO](https://help.syncfusion.com/file-formats/docio/overview)** without any dependency on Microsoft Word or Microsoft Office Interop.

---

## Overview

This project showcases common Word document generation and manipulation scenarios using the `Syncfusion.DocIO.Net.Core` library. It initializes a new Word document, configures page setup, and demonstrates key document-building features across dedicated sections before saving the output to a `.docx` file.

### Key Features Demonstrated

1. **Paragraphs & Text Formatting**
   - Inline character formatting: font family, font size, bold, italic, single underline, custom text colors (RGB), and highlight color.
   - Paragraph alignment: Left, Center, Right, and Justified.
   - Spacing and indentation: Before/After paragraph spacing, line spacing, left indent, right indent, and first-line indent.
   - Application of built-in styles (`BuiltinStyle.Heading1`).

2. **Lists**
   - **Bulleted Lists**: Default bullet style (`ApplyDefBulletStyle()`).
   - **Numbered Lists**: Default numbered list style (`ApplyDefNumberedStyle()`).
   - **Multi-Level Lists**: Nested numbering using `ListLevelNumber` across multiple hierarchy levels.

3. **Tables**
   - Dynamic table generation with defined row and column counts (`ResetCells`).
   - Custom border configuration (single border style, border color, line width).
   - Column width specification.
   - Header row designation (`IsHeader = True`) with background fill color, bold white text, and middle vertical alignment.
   - Alternating row styling (zebra striping) and cell content alignments (left-aligned text, right-aligned monetary values).

4. **Bookmarks**
   - Bookmark creation with start and end boundary markers (`AppendBookmarkStart` and `AppendBookmarkEnd`).
   - Navigating to existing bookmarks programmatically using `BookmarksNavigator`.
   - Dynamic runtime content insertion at the bookmark location (`InsertParagraphItem`).

5. **Hyperlinks**
   - **Web Links**: External website URLs (`HyperlinkType.WebLink`).
   - **Email Links**: `mailto:` links with display text (`HyperlinkType.EMailLink`).
   - **Internal Bookmark Links**: Document navigation jumping to bookmarks within the same document (`HyperlinkType.Bookmark`).

6. **Shapes & Text Boxes**
   - **Text Box (`WTextBox`)**: Configured dimensions, background fill color, border line color, and border line width with child text content.
   - **AutoShapes (`Shape`)**: Predefined shapes such as `AutoShapeType.RoundedRectangle` with custom fill color, border line weight, outline color, and embedded text body.

7. **Images**
   - Dynamic in-memory generation of PNG image bytes using pure managed code (no GDI+ or System.Drawing native dependencies).
   - Writing image to file stream and inserting it into the document (`AppendPicture`).
   - Image sizing (custom width and height).
   - Positioning, alignment (centered), and alternative text (`AlternativeText`).

---

## Project Structure

```text
word-document-generation-docio-vb-net/
├── README.md                                                    # Project documentation
└── VB.NET/
    ├── word-document-generation-docio-vb-net.slnx               # Visual Studio Solution file
    └── word-document-generation-docio-vb-net/
        ├── Program.vb                                           # Main application logic and section builders
        ├── word-document-generation-docio-vb-net.vbproj         # VB.NET project configuration (.NET 10)
        └── Output/                                              # Default output folder (generated on run)
            └── GeneratedDocument.docx                           # Generated Word document
```

---

## Prerequisites

- **[.NET SDK](https://dotnet.microsoft.com/download)** (Targeting .NET 10.0 or compatible runtime)
- **Visual Studio 2022+** or **VS Code** with the .NET SDK installed

---

## NuGet Dependencies

- **[Syncfusion.DocIO.Net.Core](https://www.nuget.org/packages/Syncfusion.DocIO.Net.Core)**: Core library for creating and manipulating Word documents in .NET.

---

## How to Run

1. Open a terminal and navigate to the project directory:
   ```bash
   cd VB.NET/word-document-generation-docio-vb-net
   ```

2. Restore packages and build the project:
   ```bash
   dotnet build
   ```

3. Run the application:
   ```bash
   dotnet run
   ```

4. Upon execution, the console displays progress for each section, and the generated Word document is saved to:
   ```text
   VB.NET/word-document-generation-docio-vb-net/Output/GeneratedDocument.docx
   ```
