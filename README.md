# utl-creating-microsoft-word-document-and-extracting-heading-and-footnotes-using-poweshell
Creating microsoft word document and extracting heading and footnotes using poweshell
    %let pgm=utl-creating-microsoft-word-document-and-extracting-heading-and-footnotes-using-poweshell;

    %stop_submission;

    Creating microsoft word document and extracting heading and footnotes using poweshell

    github
    https://tinyurl.com/3b547nvh
    https://github.com/rogerjdeangelis/utl-creating-microsoft-word-document-and-extracting-heading-and-footnotes-using-poweshell

    communities.sas
    https://tinyurl.com/yc2yyydc
    https://communities.sas.com/t5/SAS-Programming/Error-DDE-session-not-ready-ERROR-Critical-YP-Subsystem-error/m-p/818855#M323246

    Better than DDE?

    Word dynamically creates the list numbers..
    As far as I can tell you need separate code to get the numbers.
    You have to run word in the background?
    I am a real novice powershell programmer.


    AI Query (deepSeek)
    please provide a reproducible example to create and extracting headings and footnotes from a single page
    ms word document with numbered headings levels and one footnote using powershell

    IT DID NOT WORK BUT HELPED


    /********************************************************************************************************************************************/
    /*       INPUT                                          |              PROCESS                             |           OUTPUT               */
    /*       =====                                          |              =======                             |           ======               */
    /*                                                      |                                                  |                                */
    /* MICROSOFT WORD DOCUMENT                              | POWERSHELL                                       | SAS LOG                        */
    /* =======================                              |                                                  |                                */
    /*                                                      |                                                  |                                */
    /* d:/doc/sample_with_footnote.docx                     | Create a word document with headings and a       | First List Number: 1.          */
    /*                                                      | footnote. Save the doc.                          | First List Text: First item    */
    /*  1. First Item                                       | Open and extract the first list number and       | First Footnote: Footnote text. */
    /*  2. Second Item                                      | the first footnote                               |                                */
    /*                                                      |                                                  |                                */
    /*  This is a sample sentence with a footnote reference.|                                                  |                                */
    /*                                                      |                                                  |                                */
    /*  1                                                   |                                                  |                                */
    /*    This is the footnote text.                        |                                                  |                                */
    /*                                                      |                                                  |                                */
    /*---------------------------------------------------------------------------------------------------------|                                */
    /*                                                      |                                                  |                                */
    /* %utlfkil(d:/doc/sample_with_footnote.docx);          | %utl_psbegin;                                    |                                */
    /*                                                      | parmcards4;                                      |                                */
    /* %utl_psbegin;                                        | # Path to your Word document                     |                                */
    /* parmcards4;                                          | $docPath = "d:/doc/sample_with_footnote.docx"    |                                */
    /* # Create a new Word document                         |                                                  |                                */
    /* $word = New-Object -ComObject Word.Application       | # Start Word application                         |                                */
    /* $word.Visible = $true                                | $word = New-Object -ComObject Word.Application   |                                */
    /* $doc = $word.Documents.Add()                         | $word.Visible = $false                           |                                */
    /* $listTemplate =$word.ListGalleries[                  |                                                  |                                */
    /*   [Microsoft.Office.Interop.Word.WdListGalleryType]::| # Open the document                              |                                */
    /*   wdNumberGallery].ListTemplates(1)                  | $doc = $word.Documents.Open($docPath)            |                                */
    /* $selection = $word.Selection                         |                                                  |                                */
    /* # Add some paragraphs                                | foreach ($para in $doc.Paragraphs) {             |                                */
    /* $selection.TypeText("First item")                    |     $text = $para.Range.Text.Trim()              |                                */
    /* $selection.TypeParagraph()                           |     break                                        |                                */
    /* $selection.TypeText("Second item")                   | }                                                |                                */
    /* $selection.TypeParagraph()                           |                                                  |                                */
    /* $para = $doc.Paragraphs                              | # 2. Extract the first footnote (if any)         |                                */
    /* $para.Item(1).Range.ListFormat.ApplyNumberDefault()  | $firstFootnote = $null                           |                                */
    /* $para.Item(2).Range.ListFormat.ApplyNumberDefault()  | if ($doc.Footnotes.Count -ge 1) {                |                                */
    /*                                                      |   $firstFootnote=                                |                                */
    /* # Add body text with a footnote                      |      $doc.Footnotes.Item(1).Range.Text.Trim()    |                                */
    /* $para1 = $doc.Paragraphs.Add()                       | }                                                |                                */
    /* $para1.Range.Text = "see footnote"                   | $listNumber = $para.Range.ListFormat.ListString  |                                */
    /* $para1.Range.InsertParagraphAfter()                  | foreach ($para in $doc.Paragraphs) {             |                                */
    /*                                                      |    $listNumber=$para.Range.ListFormat.ListString |                                */
    /* $range = $para1.Range                                |    break                                         |                                */
    /* $range.Collapse(0)  # 0 = wdCollapseEnd              | }                                                |                                */
    /* $footnote = $doc.Footnotes.Add(                      |                                                  |                                */
    /*     $range                                           |                                                  |                                */
    /*    ,""                                               | # Output the results                             |                                */
    /*    ,"Footnote text.")                                | Write-Host "First List Number: $listNumber"      |                                */
    /*                                                      | Write-Host "First List Text: $text"              |                                */
    /* $doc.SaveAs([ref]"d:/doc/sample_with_footnote.docx") | Write-Host "First Footnote: $firstFootnote"      |                                */
    /* $doc.Close()                                         |                                                  |                                */
    /* $word.Quit()                                         | # Cleanup                                        |                                */
    /* [System.Runtime.Interopservices.Marshal]::           | $doc.Close()                                     |                                */
    /*   ReleaseComObject($doc) | Out-Null                  | $word.Quit()                                     |                                */
    /* [System.Runtime.Interopservices.Marshal]::           | [System.Runtime.Interopservices.Marshal]::       |                                */
    /*   ReleaseComObject($word) | Out-Null                 |  ReleaseComObject($word) | Out-Null              |                                */
    /* ;;;;                                                 | ;;;;                                             |                                */
    /* %utl_psend;                                          | %utl_psend;                                      |                                */
    /********************************************************************************************************************************************/

     /*                   _
    (_)_ __  _ __  _   _| |_
    | | `_ \| `_ \| | | | __|
    | | | | | |_) | |_| | |_
    |_|_| |_| .__/ \__,_|\__|
            |_|
    */

    %utlfkil(d:/doc/sample_with_footnote.docx);

    %utl_psbegin;
    parmcards4;
    # Create a new Word document
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    $doc = $word.Documents.Add()
    $listTemplate =$word.ListGalleries[
      [Microsoft.Office.Interop.Word.WdListGalleryType]::
      wdNumberGallery].ListTemplates(1)
    $selection = $word.Selection
    # Add some paragraphs
    $selection.TypeText("First item")
    $selection.TypeParagraph()
    $selection.TypeText("Second item")
    $selection.TypeParagraph()
    $para = $doc.Paragraphs
    $para.Item(1).Range.ListFormat.ApplyNumberDefault()
    $para.Item(2).Range.ListFormat.ApplyNumberDefault()

    # Add body text with a footnote
    $para1 = $doc.Paragraphs.Add()
    $para1.Range.Text = "see footnote"
    $para1.Range.InsertParagraphAfter()

    $range = $para1.Range
    $range.Collapse(0)  # 0 = wdCollapseEnd
    $footnote = $doc.Footnotes.Add(
        $range
       ,""
       ,"Footnote text.")

    $doc.SaveAs([ref]"d:/doc/sample_with_footnote.docx")
    $doc.Close()
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::
      ReleaseComObject($doc) | Out-Null
    [System.Runtime.Interopservices.Marshal]::
      ReleaseComObject($word) | Out-Null
    ;;;;
    %utl_psend;

    /**************************************************************************************************************************/
    /* d:/doc/sample_with_footnote.docx                                                                                       */
    /*                                                                                                                        */
    /*  1. First Item                                                                                                         */
    /*  2. Second Item                                                                                                        */
    /*                                                      1                                                                 */
    /*  This is a sample sentence with a footnote reference.                                                                  */
    /*                                                                                                                        */
    /*  ...                                                                                                                   */
    /*  1                                                                                                                     */
    /*  ...                                                                                                                   */
    /*    This is the footnote text.                                                                                          */
    /**************************************************************************************************************************/

    /*
     _ __  _ __ ___   ___ ___  ___ ___
    | `_ \| `__/ _ \ / __/ _ \/ __/ __|
    | |_) | | | (_) | (_|  __/\__ \__ \
    | .__/|_|  \___/ \___\___||___/___/
    |_|
    */

    %utl_psbegin;
    parmcards4;
    # Path to your Word document
    $docPath = "d:/doc/sample_with_footnote.docx"

    # Start Word application
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false

    # Open the document
    $doc = $word.Documents.Open($docPath)

    foreach ($para in $doc.Paragraphs) {
        $text = $para.Range.Text.Trim()
        break
    }

    # 2. Extract the first footnote (if any)
    $firstFootnote = $null
    if ($doc.Footnotes.Count -ge 1) {
      $firstFootnote=
         $doc.Footnotes.Item(1).Range.Text.Trim()
    }
    $listNumber = $para.Range.ListFormat.ListString
    foreach ($para in $doc.Paragraphs) {
       $listNumber=$para.Range.ListFormat.ListString
       break
    }


    # Output the results
    Write-Host "First List Number: $listNumber"
    Write-Host "First List Text: $text"
    Write-Host "First Footnote: $firstFootnote"

    # Cleanup
    $doc.Close()
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::
     ReleaseComObject($word) | Out-Null
    ;;;;
    %utl_psend;

    /**************************************************************************************************************************/
    /* First List Number: 1.                                                                                                  */
    /* First List Text: First item                                                                                            */
    /* First Footnote: Footnote text.                                                                                         */
    /**************************************************************************************************************************/

    /*              _
      ___ _ __   __| |
     / _ \ `_ \ / _` |
    |  __/ | | | (_| |
     \___|_| |_|\__,_|

    */
