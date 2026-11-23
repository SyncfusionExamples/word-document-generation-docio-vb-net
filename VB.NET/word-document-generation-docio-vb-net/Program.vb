Imports System
Imports System.IO
Imports Syncfusion.DocIO
Imports Syncfusion.DocIO.DLS
Imports Syncfusion.Drawing

Module Program

    Sub Main(args As String())
        Console.WriteLine("=================================================================")
        Console.WriteLine(" VB.NET Document Generation Using Syncfusion DocIO")
        Console.WriteLine(" Pure .NET Solution - No Microsoft Word / Interop Dependency")
        Console.WriteLine("=================================================================")
        Console.WriteLine()

        Try
            ' Create a new Word document instance
            Using document As New WordDocument()
                ' Add the primary section to the document
                Dim section As IWSection = document.AddSection()

                ' Configure standard page setup (Letter, 1-inch margins)
                section.PageSetup.PageSize = PageSize.Letter
                section.PageSetup.Margins.All = 72.0F

                ' 1. Paragraphs with Different Formatting
                AddParagraphsWithDifferentFormatting(section)

                ' 2. Lists
                AddLists(section)

                ' 3. Tables
                AddTables(section)

                ' 4. Bookmarks
                AddBookmarks(section, document)

                ' 5. Hyperlinks
                AddHyperlinks(section)

                ' 6. Shapes
                AddShapes(section)

                ' 7. Images
                AddImages(section)

                ' Save the generated document into the project's Output folder
                Dim projectRoot As String = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", ".."))
                Dim outputDir As String = Path.Combine(projectRoot, "Output")
                If Not Directory.Exists(outputDir) Then
                    Directory.CreateDirectory(outputDir)
                End If

                Dim outputPath As String = Path.Combine(outputDir, "GeneratedDocument.docx")
                Using fileStream As New FileStream(outputPath, FileMode.Create, FileAccess.Write)
                    document.Save(fileStream, FormatType.Docx)
                End Using

                Console.WriteLine()
                Console.WriteLine("SUCCESS: Document generated successfully at:")
                Console.WriteLine(outputPath)
            End Using

        Catch ex As Exception
            Console.WriteLine()
            Console.WriteLine("ERROR generating document: " & ex.Message)
            Console.WriteLine(ex.StackTrace)
        End Try

        Console.WriteLine()
        Console.WriteLine("Press any key or exit...")
    End Sub

    ''' <summary>
    ''' Feature: Paragraphs with Different Formatting
    ''' Purpose: Demonstrates various text formatting options (bold, italic, underlined,
    ''' different font sizes, colors, highlight) and paragraph formatting options
    ''' (left, center, right, justified alignments, indentation, line and paragraph spacing).
    ''' </summary>
    Public Sub AddParagraphsWithDifferentFormatting(section As IWSection)
        Console.WriteLine("[Section 1] Adding Paragraphs with Different Formatting...")

        ' Section Heading
        Dim heading As IWParagraph = section.AddParagraph()
        heading.ApplyStyle(BuiltinStyle.Heading1)
        heading.ParagraphFormat.BeforeSpacing = 12.0F
        heading.ParagraphFormat.AfterSpacing = 6.0F
        heading.AppendText("Paragraphs with Different Formatting")

        ' Paragraph demonstrating various character/font formatting options
        Dim pTextFormatting As IWParagraph = section.AddParagraph()
        pTextFormatting.ParagraphFormat.BeforeSpacing = 4.0F
        pTextFormatting.ParagraphFormat.AfterSpacing = 6.0F
        pTextFormatting.ParagraphFormat.LineSpacing = 14.0F

        Dim rIntro As IWTextRange = pTextFormatting.AppendText("This paragraph illustrates inline font styling: ")
        rIntro.CharacterFormat.FontName = "Calibri"
        rIntro.CharacterFormat.FontSize = 11.0F

        Dim rBold As IWTextRange = pTextFormatting.AppendText("bold text, ")
        rBold.CharacterFormat.FontName = "Calibri"
        rBold.CharacterFormat.FontSize = 11.0F
        rBold.CharacterFormat.Bold = True

        Dim rItalic As IWTextRange = pTextFormatting.AppendText("italic text, ")
        rItalic.CharacterFormat.FontName = "Calibri"
        rItalic.CharacterFormat.FontSize = 11.0F
        rItalic.CharacterFormat.Italic = True

        Dim rUnderline As IWTextRange = pTextFormatting.AppendText("underlined text, ")
        rUnderline.CharacterFormat.FontName = "Calibri"
        rUnderline.CharacterFormat.FontSize = 11.0F
        rUnderline.CharacterFormat.UnderlineStyle = UnderlineStyle.Single

        Dim rSizeSmall As IWTextRange = pTextFormatting.AppendText("small 9pt text, ")
        rSizeSmall.CharacterFormat.FontName = "Calibri"
        rSizeSmall.CharacterFormat.FontSize = 9.0F

        Dim rSizeLarge As IWTextRange = pTextFormatting.AppendText("large 14pt text, ")
        rSizeLarge.CharacterFormat.FontName = "Calibri"
        rSizeLarge.CharacterFormat.FontSize = 14.0F

        Dim rColorNavy As IWTextRange = pTextFormatting.AppendText("navy colored text, ")
        rColorNavy.CharacterFormat.FontName = "Calibri"
        rColorNavy.CharacterFormat.FontSize = 11.0F
        rColorNavy.CharacterFormat.TextColor = Color.FromArgb(15, 76, 129)

        Dim rColorCrimson As IWTextRange = pTextFormatting.AppendText("crimson colored text, ")
        rColorCrimson.CharacterFormat.FontName = "Calibri"
        rColorCrimson.CharacterFormat.FontSize = 11.0F
        rColorCrimson.CharacterFormat.TextColor = Color.FromArgb(178, 34, 34)

        Dim rHighlight As IWTextRange = pTextFormatting.AppendText("and highlighted text.")
        rHighlight.CharacterFormat.FontName = "Calibri"
        rHighlight.CharacterFormat.FontSize = 11.0F
        rHighlight.CharacterFormat.HighlightColor = Color.Yellow

        ' Left-aligned paragraph
        Dim pLeft As IWParagraph = section.AddParagraph()
        pLeft.ParagraphFormat.BeforeSpacing = 4.0F
        pLeft.ParagraphFormat.AfterSpacing = 4.0F
        pLeft.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left
        Dim rLeft As IWTextRange = pLeft.AppendText("Left-Aligned Paragraph: This paragraph is aligned to the left margin. Standard document text typically uses left alignment for effortless readability and natural reading flow.")
        rLeft.CharacterFormat.FontName = "Calibri"
        rLeft.CharacterFormat.FontSize = 10.5F

        ' Center-aligned paragraph
        Dim pCenter As IWParagraph = section.AddParagraph()
        pCenter.ParagraphFormat.BeforeSpacing = 4.0F
        pCenter.ParagraphFormat.AfterSpacing = 4.0F
        pCenter.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center
        Dim rCenter As IWTextRange = pCenter.AppendText("Center-Aligned Paragraph: Useful for titles, section subtitles, and callout statements.")
        rCenter.CharacterFormat.FontName = "Calibri"
        rCenter.CharacterFormat.FontSize = 10.5F
        rCenter.CharacterFormat.Italic = True

        ' Right-aligned paragraph
        Dim pRight As IWParagraph = section.AddParagraph()
        pRight.ParagraphFormat.BeforeSpacing = 4.0F
        pRight.ParagraphFormat.AfterSpacing = 4.0F
        pRight.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right
        Dim rRight As IWTextRange = pRight.AppendText("Right-Aligned Paragraph: Document Reference: DOC-REF-2026-001 | Date: November 2026")
        rRight.CharacterFormat.FontName = "Calibri"
        rRight.CharacterFormat.FontSize = 9.5F
        rRight.CharacterFormat.TextColor = Color.FromArgb(100, 100, 100)

        ' Justified paragraph
        Dim pJustified As IWParagraph = section.AddParagraph()
        pJustified.ParagraphFormat.BeforeSpacing = 4.0F
        pJustified.ParagraphFormat.AfterSpacing = 6.0F
        pJustified.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Justify
        Dim rJustified As IWTextRange = pJustified.AppendText("Justified Paragraph: Syncfusion DocIO is a comprehensive .NET Word library that allows developers to create, read, edit, and convert Microsoft Word documents programmatically. The text in this paragraph is justified evenly across both left and right margins, producing clean vertical edges along each side of the column.")
        rJustified.CharacterFormat.FontName = "Calibri"
        rJustified.CharacterFormat.FontSize = 10.5F

        ' Paragraph indentation and spacing
        Dim pIndent As IWParagraph = section.AddParagraph()
        pIndent.ParagraphFormat.BeforeSpacing = 6.0F
        pIndent.ParagraphFormat.AfterSpacing = 10.0F
        pIndent.ParagraphFormat.LeftIndent = 36.0F
        pIndent.ParagraphFormat.RightIndent = 36.0F
        pIndent.ParagraphFormat.FirstLineIndent = 18.0F
        pIndent.ParagraphFormat.LineSpacing = 16.0F
        Dim rIndent As IWTextRange = pIndent.AppendText("Indented Paragraph: This paragraph demonstrates custom left indent (36 pt), right indent (36 pt), a first-line indent (18 pt), and custom line spacing (16 pt). Such formatting is commonly applied to quotations, legal clauses, and highlighted advisory blocks.")
        rIndent.CharacterFormat.FontName = "Georgia"
        rIndent.CharacterFormat.FontSize = 10.0F
        rIndent.CharacterFormat.TextColor = Color.FromArgb(40, 40, 40)
    End Sub

    ''' <summary>
    ''' Feature: Lists
    ''' Purpose: Demonstrates bulleted lists, numbered lists, and multi-level lists.
    ''' </summary>
    Public Sub AddLists(section As IWSection)
        Console.WriteLine("[Section 2] Adding Lists...")

        ' Section Heading
        Dim heading As IWParagraph = section.AddParagraph()
        heading.ApplyStyle(BuiltinStyle.Heading1)
        heading.ParagraphFormat.BeforeSpacing = 14.0F
        heading.ParagraphFormat.AfterSpacing = 6.0F
        heading.AppendText("Lists")

        ' Subheading - Bulleted Lists
        Dim pSub1 As IWParagraph = section.AddParagraph()
        pSub1.ApplyStyle(BuiltinStyle.Heading2)
        pSub1.ParagraphFormat.BeforeSpacing = 6.0F
        pSub1.ParagraphFormat.AfterSpacing = 4.0F
        pSub1.AppendText("Bulleted List")

        Dim bulletItems As String() = {
            "Comprehensive Word processing without Microsoft Office dependencies.",
            "High performance document generation and manipulation in server environments.",
            "Full support for DOCX, DOC, RTF, WordML, HTML, and PDF conversion."
        }
        For Each item As String In bulletItems
            Dim pBullet As IWParagraph = section.AddParagraph()
            pBullet.ListFormat.ApplyDefBulletStyle()
            pBullet.ParagraphFormat.BeforeSpacing = 2.0F
            pBullet.ParagraphFormat.AfterSpacing = 2.0F
            Dim r As IWTextRange = pBullet.AppendText(item)
            r.CharacterFormat.FontName = "Calibri"
            r.CharacterFormat.FontSize = 10.5F
        Next

        ' Subheading - Numbered Lists
        Dim pSub2 As IWParagraph = section.AddParagraph()
        pSub2.ApplyStyle(BuiltinStyle.Heading2)
        pSub2.ParagraphFormat.BeforeSpacing = 8.0F
        pSub2.ParagraphFormat.AfterSpacing = 4.0F
        pSub2.AppendText("Numbered List")

        Dim numberedItems As String() = {
            "Initialize a new WordDocument instance.",
            "Add one or more sections with customized page setup.",
            "Populate document elements such as paragraphs, tables, shapes, and images.",
            "Save the document stream to a file or client response."
        }
        For Each item As String In numberedItems
            Dim pNumber As IWParagraph = section.AddParagraph()
            pNumber.ListFormat.ApplyDefNumberedStyle()
            pNumber.ParagraphFormat.BeforeSpacing = 2.0F
            pNumber.ParagraphFormat.AfterSpacing = 2.0F
            Dim r As IWTextRange = pNumber.AppendText(item)
            r.CharacterFormat.FontName = "Calibri"
            r.CharacterFormat.FontSize = 10.5F
        Next

        ' Subheading - Multi-level Lists
        Dim pSub3 As IWParagraph = section.AddParagraph()
        pSub3.ApplyStyle(BuiltinStyle.Heading2)
        pSub3.ParagraphFormat.BeforeSpacing = 8.0F
        pSub3.ParagraphFormat.AfterSpacing = 4.0F
        pSub3.AppendText("Multi-Level List")

        Dim multiLevelData As Tuple(Of Integer, String)() = {
            Tuple.Create(0, "Document Processing Components"),
            Tuple.Create(1, "Syncfusion DocIO (.NET Word library)"),
            Tuple.Create(1, "Syncfusion XlsIO (.NET Excel library)"),
            Tuple.Create(1, "Syncfusion Presentation (.NET PowerPoint library)"),
            Tuple.Create(0, "Supported Target Frameworks"),
            Tuple.Create(1, ".NET 6, .NET 8, .NET 9, and .NET 10"),
            Tuple.Create(1, ".NET Standard 2.0 and .NET Framework 4.6.2+")
        }

        For Each entry As Tuple(Of Integer, String) In multiLevelData
            Dim pMulti As IWParagraph = section.AddParagraph()
            pMulti.ListFormat.ApplyDefNumberedStyle()
            pMulti.ListFormat.ListLevelNumber = entry.Item1
            pMulti.ParagraphFormat.BeforeSpacing = 2.0F
            pMulti.ParagraphFormat.AfterSpacing = 2.0F
            Dim r As IWTextRange = pMulti.AppendText(entry.Item2)
            r.CharacterFormat.FontName = "Calibri"
            r.CharacterFormat.FontSize = 10.5F
        Next
    End Sub

    ''' <summary>
    ''' Feature: Tables
    ''' Purpose: Demonstrates creating tables, adding rows and columns, populating with data,
    ''' applying single borders, cell background colors, formatting text, column widths, and cell alignment.
    ''' </summary>
    Public Sub AddTables(section As IWSection)
        Console.WriteLine("[Section 3] Adding Tables...")

        ' Section Heading
        Dim heading As IWParagraph = section.AddParagraph()
        heading.ApplyStyle(BuiltinStyle.Heading1)
        heading.ParagraphFormat.BeforeSpacing = 14.0F
        heading.ParagraphFormat.AfterSpacing = 6.0F
        heading.AppendText("Tables")

        Dim pIntro As IWParagraph = section.AddParagraph()
        pIntro.ParagraphFormat.AfterSpacing = 6.0F
        Dim rIntro As IWTextRange = pIntro.AppendText("The table below demonstrates custom borders, navy header shading, alternating row background colors, explicit column widths, and numerical right alignment.")
        rIntro.CharacterFormat.FontName = "Calibri"
        rIntro.CharacterFormat.FontSize = 10.5F

        ' Create table dynamically (4 data rows + 1 header = 5 rows, 4 columns)
        Dim table As WTable = TryCast(section.AddTable(), WTable)
        If table Is Nothing Then Return

        Dim rowCount As Integer = 5
        Dim colCount As Integer = 4
        table.ResetCells(rowCount, colCount)

        ' Apply table borders
        table.TableFormat.Borders.BorderType = BorderStyle.Single
        table.TableFormat.Borders.Color = Color.FromArgb(200, 205, 215)
        table.TableFormat.Borders.LineWidth = 0.75F
        table.TableFormat.IsAutoResized = False

        ' Set column widths (Total: 468 pt)
        Dim columnWidths As Single() = {70.0F, 198.0F, 100.0F, 100.0F}
        For Each row As WTableRow In table.Rows
            For c As Integer = 0 To columnWidths.Length - 1
                row.Cells(c).Width = columnWidths(c)
            Next
        Next

        ' Format and populate Header Row
        Dim headerRow As WTableRow = table.Rows(0)
        headerRow.IsHeader = True
        Dim headers As String() = {"ID", "Module Name", "Platform", "Annual Cost"}
        For c As Integer = 0 To headers.Length - 1
            Dim cell As WTableCell = headerRow.Cells(c)
            cell.CellFormat.BackColor = Color.FromArgb(15, 76, 129)
            cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle

            Dim p As IWParagraph = cell.AddParagraph()
            p.ParagraphFormat.BeforeSpacing = 5.0F
            p.ParagraphFormat.AfterSpacing = 5.0F
            p.ParagraphFormat.HorizontalAlignment = If(c = 3, HorizontalAlignment.Right, HorizontalAlignment.Left)

            Dim r As IWTextRange = p.AppendText(headers(c))
            r.CharacterFormat.FontName = "Calibri"
            r.CharacterFormat.FontSize = 10.0F
            r.CharacterFormat.Bold = True
            r.CharacterFormat.TextColor = Color.White
        Next

        ' Sample data rows
        Dim tableData As String(,) = {
            {"MOD-101", "Word Processing (DocIO)", "Cross-Platform", "$1,499.00"},
            {"MOD-102", "Spreadsheet Engine (XlsIO)", "Cross-Platform", "$1,499.00"},
            {"MOD-103", "PDF Generator & Viewer", "Cross-Platform", "$1,799.00"},
            {"MOD-104", "Presentation Graphics Engine", "Cross-Platform", "$1,299.00"}
        }

        For r As Integer = 0 To 3
            Dim row As WTableRow = table.Rows(r + 1)
            Dim isEven As Boolean = (r Mod 2 = 0)
            Dim bg As Color = If(isEven, Color.White, Color.FromArgb(245, 248, 252))

            For c As Integer = 0 To colCount - 1
                Dim cell As WTableCell = row.Cells(c)
                cell.CellFormat.BackColor = bg
                cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle

                Dim p As IWParagraph = cell.AddParagraph()
                p.ParagraphFormat.BeforeSpacing = 4.0F
                p.ParagraphFormat.AfterSpacing = 4.0F
                p.ParagraphFormat.HorizontalAlignment = If(c = 3, HorizontalAlignment.Right, HorizontalAlignment.Left)

                Dim rText As IWTextRange = p.AppendText(tableData(r, c))
                rText.CharacterFormat.FontName = "Calibri"
                rText.CharacterFormat.FontSize = 9.5F
                rText.CharacterFormat.TextColor = Color.FromArgb(40, 40, 40)
            Next
        Next
    End Sub

    ''' <summary>
    ''' Feature: Bookmarks
    ''' Purpose: Demonstrates creating bookmarks, navigating to them with BookmarksNavigator,
    ''' and inserting dynamic content at bookmark locations.
    ''' </summary>
    Public Sub AddBookmarks(section As IWSection, document As WordDocument)
        Console.WriteLine("[Section 4] Adding Bookmarks...")

        ' Section Heading
        Dim heading As IWParagraph = section.AddParagraph()
        heading.ApplyStyle(BuiltinStyle.Heading1)
        heading.ParagraphFormat.BeforeSpacing = 14.0F
        heading.ParagraphFormat.AfterSpacing = 6.0F
        heading.AppendText("Bookmarks")

        ' Paragraph hosting the bookmark
        Dim pBookmarkHost As IWParagraph = section.AddParagraph()
        pBookmarkHost.ParagraphFormat.BeforeSpacing = 4.0F
        pBookmarkHost.ParagraphFormat.AfterSpacing = 4.0F

        Dim rPrefix As IWTextRange = pBookmarkHost.AppendText("Dynamic Bookmark Slot: [ ")
        rPrefix.CharacterFormat.FontName = "Calibri"
        rPrefix.CharacterFormat.FontSize = 10.5F

        ' Create Bookmark: "SyncfusionDocIOBookmark"
        Const bookmarkName As String = "SyncfusionDocIOBookmark"
        pBookmarkHost.AppendBookmarkStart(bookmarkName)
        pBookmarkHost.AppendBookmarkEnd(bookmarkName)

        Dim rSuffix As IWTextRange = pBookmarkHost.AppendText(" ]")
        rSuffix.CharacterFormat.FontName = "Calibri"
        rSuffix.CharacterFormat.FontSize = 10.5F

        ' Navigate to the bookmark and insert dynamic content
        Dim navigator As New BookmarksNavigator(document)
        navigator.MoveToBookmark(bookmarkName)

        ' Insert styled dynamic text inside the bookmark via BookmarksNavigator
        Dim dynamicRange As WTextRange = TryCast(navigator.InsertParagraphItem(ParagraphItemType.TextRange), WTextRange)
        If dynamicRange IsNot Nothing Then
            dynamicRange.Text = "Dynamic Content Inserted via BookmarksNavigator at Runtime"
            dynamicRange.CharacterFormat.FontName = "Calibri"
            dynamicRange.CharacterFormat.FontSize = 10.5F
            dynamicRange.CharacterFormat.Bold = True
            dynamicRange.CharacterFormat.TextColor = Color.FromArgb(15, 76, 129)
        End If

        ' Explanatory note
        Dim pNote As IWParagraph = section.AddParagraph()
        pNote.ParagraphFormat.BeforeSpacing = 2.0F
        pNote.ParagraphFormat.AfterSpacing = 6.0F
        Dim rNote As IWTextRange = pNote.AppendText("The text inside brackets above was injected dynamically using BookmarksNavigator.MoveToBookmark() and InsertParagraphItem().")
        rNote.CharacterFormat.FontName = "Calibri"
        rNote.CharacterFormat.FontSize = 9.5F
        rNote.CharacterFormat.Italic = True
        rNote.CharacterFormat.TextColor = Color.FromArgb(100, 100, 100)
    End Sub

    ''' <summary>
    ''' Feature: Hyperlinks
    ''' Purpose: Demonstrates website hyperlinks, email hyperlinks, and internal
    ''' document hyperlinks navigating to bookmarks.
    ''' </summary>
    Public Sub AddHyperlinks(section As IWSection)
        Console.WriteLine("[Section 5] Adding Hyperlinks...")

        ' Section Heading
        Dim heading As IWParagraph = section.AddParagraph()
        heading.ApplyStyle(BuiltinStyle.Heading1)
        heading.ParagraphFormat.BeforeSpacing = 14.0F
        heading.ParagraphFormat.AfterSpacing = 6.0F
        heading.AppendText("Hyperlinks")

        ' 1. Website Hyperlink
        Dim pWeb As IWParagraph = section.AddParagraph()
        pWeb.ParagraphFormat.BeforeSpacing = 3.0F
        pWeb.ParagraphFormat.AfterSpacing = 3.0F
        Dim rWebLabel As IWTextRange = pWeb.AppendText("1. Website Hyperlink: ")
        rWebLabel.CharacterFormat.FontName = "Calibri"
        rWebLabel.CharacterFormat.FontSize = 10.5F
        rWebLabel.CharacterFormat.Bold = True
        pWeb.AppendHyperlink("https://www.syncfusion.com", "Visit Syncfusion Official Website", HyperlinkType.WebLink)

        ' 2. Email Hyperlink
        Dim pMail As IWParagraph = section.AddParagraph()
        pMail.ParagraphFormat.BeforeSpacing = 3.0F
        pMail.ParagraphFormat.AfterSpacing = 3.0F
        Dim rMailLabel As IWTextRange = pMail.AppendText("2. Email Hyperlink: ")
        rMailLabel.CharacterFormat.FontName = "Calibri"
        rMailLabel.CharacterFormat.FontSize = 10.5F
        rMailLabel.CharacterFormat.Bold = True
        pMail.AppendHyperlink("mailto:support@syncfusion.com", "Contact Syncfusion Support", HyperlinkType.EMailLink)

        ' 3. Internal Document Hyperlink using Bookmark
        Dim pInternal As IWParagraph = section.AddParagraph()
        pInternal.ParagraphFormat.BeforeSpacing = 3.0F
        pInternal.ParagraphFormat.AfterSpacing = 6.0F
        Dim rIntLabel As IWTextRange = pInternal.AppendText("3. Internal Document Bookmark Link: ")
        rIntLabel.CharacterFormat.FontName = "Calibri"
        rIntLabel.CharacterFormat.FontSize = 10.5F
        rIntLabel.CharacterFormat.Bold = True
        pInternal.AppendHyperlink("SyncfusionDocIOBookmark", "Jump to the Bookmarks Section Above", HyperlinkType.Bookmark)
    End Sub

    ''' <summary>
    ''' Feature: Shapes
    ''' Purpose: Demonstrates inserting text boxes, adding basic shapes supported by DocIO
    ''' (e.g. RoundedRectangle), and applying fill color, outline color, and size.
    ''' </summary>
    Public Sub AddShapes(section As IWSection)
        Console.WriteLine("[Section 6] Adding Shapes...")

        ' Section Heading
        Dim heading As IWParagraph = section.AddParagraph()
        heading.ApplyStyle(BuiltinStyle.Heading1)
        heading.ParagraphFormat.BeforeSpacing = 14.0F
        heading.ParagraphFormat.AfterSpacing = 6.0F
        heading.AppendText("Shapes")

        ' Paragraph hosting the text box
        Dim pTextBoxHost As IWParagraph = section.AddParagraph()
        pTextBoxHost.ParagraphFormat.BeforeSpacing = 4.0F
        pTextBoxHost.ParagraphFormat.AfterSpacing = 6.0F

        ' 1. Insert Text Box
        Dim textBox As WTextBox = TryCast(pTextBoxHost.AppendTextBox(420.0F, 50.0F), WTextBox)
        If textBox IsNot Nothing Then
            textBox.TextBoxFormat.FillColor = Color.FromArgb(240, 245, 255)
            textBox.TextBoxFormat.LineColor = Color.FromArgb(15, 76, 129)
            textBox.TextBoxFormat.LineWidth = 1.2F

            Dim pInBox As IWParagraph = textBox.TextBoxBody.AddParagraph()
            pInBox.ParagraphFormat.BeforeSpacing = 4.0F
            pInBox.ParagraphFormat.AfterSpacing = 2.0F
            Dim rBoxTitle As IWTextRange = pInBox.AppendText("Text Box Component: ")
            rBoxTitle.CharacterFormat.FontName = "Calibri"
            rBoxTitle.CharacterFormat.FontSize = 10.5F
            rBoxTitle.CharacterFormat.Bold = True
            rBoxTitle.CharacterFormat.TextColor = Color.FromArgb(15, 76, 129)

            Dim rBoxBody As IWTextRange = pInBox.AppendText("Demonstrates text boxes styled with custom fill color, border line color, and line width.")
            rBoxBody.CharacterFormat.FontName = "Calibri"
            rBoxBody.CharacterFormat.FontSize = 9.5F
        End If

        ' Paragraph hosting basic shapes
        Dim pShapeHost As IWParagraph = section.AddParagraph()
        pShapeHost.ParagraphFormat.BeforeSpacing = 6.0F
        pShapeHost.ParagraphFormat.AfterSpacing = 6.0F

        ' 2. Add Basic Shape: Rounded Rectangle
        Dim roundedRect As Shape = TryCast(pShapeHost.AppendShape(AutoShapeType.RoundedRectangle, 200.0F, 50.0F), Shape)
        If roundedRect IsNot Nothing Then
            roundedRect.FillFormat.Color = Color.FromArgb(220, 240, 230)
            roundedRect.LineFormat.Color = Color.FromArgb(34, 139, 34)
            roundedRect.LineFormat.Weight = 1.5F

            Dim pShapeText As IWParagraph = roundedRect.TextBody.AddParagraph()
            pShapeText.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center
            pShapeText.ParagraphFormat.BeforeSpacing = 6.0F
            Dim rShapeText As IWTextRange = pShapeText.AppendText("Rounded Rectangle Shape")
            rShapeText.CharacterFormat.FontName = "Calibri"
            rShapeText.CharacterFormat.FontSize = 9.5F
            rShapeText.CharacterFormat.Bold = True
            rShapeText.CharacterFormat.TextColor = Color.FromArgb(20, 90, 20)
        End If
    End Sub

    ''' <summary>
    ''' Feature: Images
    ''' Purpose: Demonstrates creating a sample image file on disk, inserting an image
    ''' from a file path, resizing the image, and positioning it within the document.
    ''' </summary>
    Public Sub AddImages(section As IWSection)
        Console.WriteLine("[Section 7] Adding Images...")

        ' Section Heading
        Dim heading As IWParagraph = section.AddParagraph()
        heading.ApplyStyle(BuiltinStyle.Heading1)
        heading.ParagraphFormat.BeforeSpacing = 14.0F
        heading.ParagraphFormat.AfterSpacing = 6.0F
        heading.AppendText("Images")

        Dim pIntro As IWParagraph = section.AddParagraph()
        pIntro.ParagraphFormat.AfterSpacing = 6.0F
        Dim rIntro As IWTextRange = pIntro.AppendText("This section demonstrates creating and inserting an image from a file, resizing it to specified dimensions, and positioning it inline with center alignment.")
        rIntro.CharacterFormat.FontName = "Calibri"
        rIntro.CharacterFormat.FontSize = 10.5F

        ' Ensure sample image file exists on disk
        Dim tempImagePath As String = Path.Combine(Path.GetTempPath(), "SyncfusionDocIOSampleImage.png")
        Dim pngBytes As Byte() = GenerateSamplePng(240, 60, 15, 76, 129)
        File.WriteAllBytes(tempImagePath, pngBytes)

        ' Paragraph hosting the image
        Dim pImage As IWParagraph = section.AddParagraph()
        pImage.ParagraphFormat.BeforeSpacing = 6.0F
        pImage.ParagraphFormat.AfterSpacing = 6.0F
        pImage.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center

        ' Insert image from file path using FileStream (pure .NET standard I/O)
        Using imgStream As New FileStream(tempImagePath, FileMode.Open, FileAccess.Read)
            Dim picture As IWPicture = pImage.AppendPicture(imgStream)

            ' Resize image
            picture.Width = 240.0F
            picture.Height = 60.0F

            ' Positioning within the document
            picture.TextWrappingStyle = TextWrappingStyle.Inline
            picture.AlternativeText = "Syncfusion DocIO Sample Banner Image"
        End Using

        ' Clean up temporary image file
        Try
            If File.Exists(tempImagePath) Then File.Delete(tempImagePath)
        Catch
            ' Non-critical cleanup
        End Try
    End Sub

    ''' <summary>
    ''' Helper to synthesize a valid PNG image in pure managed code
    ''' without external GDI+ or System.Drawing dependencies.
    ''' </summary>
    Private Function GenerateSamplePng(width As Integer, height As Integer, r As Byte, g As Byte, b As Byte) As Byte()
        Using ms As New MemoryStream()
            ' PNG Header
            ms.Write(New Byte() {&H89, &H50, &H4E, &H47, &HD, &HA, &H1A, &HA}, 0, 8)

            ' IHDR Chunk
            Dim ihdr(12) As Byte
            ihdr(0) = CByte((width >> 24) And &HFF)
            ihdr(1) = CByte((width >> 16) And &HFF)
            ihdr(2) = CByte((width >> 8) And &HFF)
            ihdr(3) = CByte(width And &HFF)
            ihdr(4) = CByte((height >> 24) And &HFF)
            ihdr(5) = CByte((height >> 16) And &HFF)
            ihdr(6) = CByte((height >> 8) And &HFF)
            ihdr(7) = CByte(height And &HFF)
            ihdr(8) = 8   ' Bit depth
            ihdr(9) = 2   ' Color type 2 (RGB)
            ihdr(10) = 0  ' Compression
            ihdr(11) = 0  ' Filter
            ihdr(12) = 0  ' Interlace
            WritePngChunk(ms, "IHDR", ihdr)

            ' IDAT Chunk - raw scanlines with uncompressed deflate
            Dim scanlineLen As Integer = 1 + width * 3
            Dim rawData(height * scanlineLen - 1) As Byte
            Dim idx As Integer = 0
            For y As Integer = 0 To height - 1
                rawData(idx) = 0 ' Filter type 0 (None)
                idx += 1
                For x As Integer = 0 To width - 1
                    rawData(idx) = r
                    rawData(idx + 1) = g
                    rawData(idx + 2) = b
                    idx += 3
                Next
            Next

            ' Uncompressed deflate stream (RFC 1950 zlib wrapper)
            Using idatStream As New MemoryStream()
                idatStream.WriteByte(&H78) ' CMF (deflate, 32K window)
                idatStream.WriteByte(&H1)  ' FLG

                Dim offset As Integer = 0
                Dim remaining As Integer = rawData.Length
                While remaining > 0
                    Dim blockLen As Integer = Math.Min(remaining, 65535)
                    Dim isLast As Boolean = (blockLen = remaining)
                    idatStream.WriteByte(CByte(If(isLast, 1, 0)))
                    idatStream.WriteByte(CByte(blockLen And &HFF))
                    idatStream.WriteByte(CByte((blockLen >> 8) And &HFF))
                    Dim nlen As Integer = blockLen Xor &HFFFF
                    idatStream.WriteByte(CByte(nlen And &HFF))
                    idatStream.WriteByte(CByte((nlen >> 8) And &HFF))
                    idatStream.Write(rawData, offset, blockLen)
                    offset += blockLen
                    remaining -= blockLen
                End While

                ' Adler-32 checksum
                Dim a As UInteger = 1, bSum As UInteger = 0
                For Each by As Byte In rawData
                    a = (a + by) Mod 65521UI
                    bSum = (bSum + a) Mod 65521UI
                Next
                Dim adler As UInteger = (bSum << 16) Or a
                idatStream.WriteByte(CByte((adler >> 24) And &HFF))
                idatStream.WriteByte(CByte((adler >> 16) And &HFF))
                idatStream.WriteByte(CByte((adler >> 8) And &HFF))
                idatStream.WriteByte(CByte(adler And &HFF))

                WritePngChunk(ms, "IDAT", idatStream.ToArray())
            End Using

            ' IEND Chunk
            WritePngChunk(ms, "IEND", New Byte() {})

            Return ms.ToArray()
        End Using
    End Function

    Private Sub WritePngChunk(ms As Stream, chunkType As String, data As Byte())
        Dim length As Integer = data.Length
        ms.WriteByte(CByte((length >> 24) And &HFF))
        ms.WriteByte(CByte((length >> 16) And &HFF))
        ms.WriteByte(CByte((length >> 8) And &HFF))
        ms.WriteByte(CByte(length And &HFF))

        Dim typeBytes As Byte() = System.Text.Encoding.ASCII.GetBytes(chunkType)
        ms.Write(typeBytes, 0, 4)
        If data.Length > 0 Then
            ms.Write(data, 0, data.Length)
        End If

        Dim crc As UInteger = CalculateCrc32(typeBytes, data)
        ms.WriteByte(CByte((crc >> 24) And &HFF))
        ms.WriteByte(CByte((crc >> 16) And &HFF))
        ms.WriteByte(CByte((crc >> 8) And &HFF))
        ms.WriteByte(CByte(crc And &HFF))
    End Sub

    Private Function CalculateCrc32(header As Byte(), body As Byte()) As UInteger
        Dim crc As UInteger = &HFFFFFFFFUI
        For Each b As Byte In header
            crc = CrcTable((crc Xor b) And &HFF) Xor (crc >> 8)
        Next
        For Each b As Byte In body
            crc = CrcTable((crc Xor b) And &HFF) Xor (crc >> 8)
        Next
        Return crc Xor &HFFFFFFFFUI
    End Function

    Private ReadOnly CrcTable As UInteger() = InitializeCrcTable()

    Private Function InitializeCrcTable() As UInteger()
        Dim table(255) As UInteger
        For i As UInteger = 0 To 255UI
            Dim c As UInteger = i
            For j As Integer = 0 To 7
                If (c And 1UI) <> 0 Then
                    c = &HEDB88320UI Xor (c >> 1)
                Else
                    c = c >> 1
                End If
            Next
            table(i) = c
        Next
        Return table
    End Function

End Module
