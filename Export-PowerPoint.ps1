<#
.Synopsis
   Convert PPT files to PDF files using PowerPoint COM Object
.DESCRIPTION
   Convert PPT files passed as path to PDF files so you can share them afterwards
   to external audience
.EXAMPLE
   Convert-PPT2PDF -Path c:\MyPPTFiles
   Will convert files at this path
.EXAMPLE
   Convert-PPT2PDF
   Will Convert files in current folder
.INPUTS
   Path where files are stored
.OUTPUTS
   None
.NOTES
   PowerPoint must be installed on the computer to convert the files.
   Only Windows PowerShell is compatible to run this script (no PowerShell Core)
#>

[CmdletBinding( SupportsShouldProcess=$true, 
                PositionalBinding=$false,
                ConfirmImpact='Medium')]
[OutputType([void])]
Param
(
    # Path 
    [Parameter(Mandatory=$False,
    ValueFromPipeline=$True,
    ValueFromPipelinebyPropertyName=$True)]
    [PSDefaultValue(Help='Path where the files are. Default is working directory')]
    [SupportsWildcards()]
    [string[]]$Path='.'
)

Begin
{
    Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint > $null
	Add-Type -AssemblyName Office > $null
    
    $msoFalse =  [Microsoft.Office.Core.MsoTristate]::msoFalse
	$msoTrue =  [Microsoft.Office.Core.MsoTristate]::msoTrue
    
    $ppFixedFormatIntentScreen = [Microsoft.Office.Interop.PowerPoint.PpFixedFormatIntent]::ppFixedFormatIntentScreen # Intent is to view exported file on screen.
	$ppFixedFormatIntentPrint =  [Microsoft.Office.Interop.PowerPoint.PpFixedFormatIntent]::ppFixedFormatIntentPrint  # Intent is to print exported file.

	$ppFixedFormatTypeXPS = [Microsoft.Office.Interop.PowerPoint.PpFixedFormatType]::ppFixedFormatTypeXPS  # XPS format
	$ppFixedFormatTypePDF = [Microsoft.Office.Interop.PowerPoint.PpFixedFormatType]::ppFixedFormatTypePDF  # PDF format

	$ppPrintHandoutVerticalFirst   = [Microsoft.Office.Interop.PowerPoint.PpPrintHandoutOrder]::ppPrintHandoutVerticalFirst   # Slides are ordered vertically, with the first slide in the upper-left corner and the second slide below it.
	$ppPrintHandoutHorizontalFirst = [Microsoft.Office.Interop.PowerPoint.PpPrintHandoutOrder]::ppPrintHandoutHorizontalFirst # Slides are ordered horizontally, with the first slide in the upper-left corner and the second slide to the right of it.

	$ppPrintOutputSlides             = [Microsoft.Office.Interop.PowerPoint.PpPrintOutputType]::ppPrintOutputSlides              # Slides
	$ppPrintOutputTwoSlideHandouts   = [Microsoft.Office.Interop.PowerPoint.PpPrintOutputType]::ppPrintOutputTwoSlideHandouts    # Two Slide Handouts
	$ppPrintOutputThreeSlideHandouts = [Microsoft.Office.Interop.PowerPoint.PpPrintOutputType]::ppPrintOutputThreeSlideHandouts  # Three Slide Handouts
	$ppPrintOutputSixSlideHandouts   = [Microsoft.Office.Interop.PowerPoint.PpPrintOutputType]::ppPrintOutputSixSlideHandouts    # Six Slide Handouts
	$ppPrintOutputNotesPages         = [Microsoft.Office.Interop.PowerPoint.PpPrintOutputType]::ppPrintOutputNotesPages          # Notes Pages
	$ppPrintOutputOutline            = [Microsoft.Office.Interop.PowerPoint.PpPrintOutputType]::ppPrintOutputOutline             # Outline
	$ppPrintOutputBuildSlides        = [Microsoft.Office.Interop.PowerPoint.PpPrintOutputType]::ppPrintOutputBuildSlides         # Build Slides
	$ppPrintOutputFourSlideHandouts  = [Microsoft.Office.Interop.PowerPoint.PpPrintOutputType]::ppPrintOutputFourSlideHandouts   # Four Slide Handouts
	$ppPrintOutputNineSlideHandouts  = [Microsoft.Office.Interop.PowerPoint.PpPrintOutputType]::ppPrintOutputNineSlideHandouts   # Nine Slide Handouts
	$ppPrintOutputOneSlideHandouts   = [Microsoft.Office.Interop.PowerPoint.PpPrintOutputType]::ppPrintOutputOneSlideHandouts    # Single Slide Handouts

	$ppPrintAll            = [Microsoft.Office.Interop.PowerPoint.PpPrintRangeType]::ppPrintAll            # Print all slides in the presentation.
	$ppPrintSelection      = [Microsoft.Office.Interop.PowerPoint.PpPrintRangeType]::ppPrintSelection      # Print a selection of slides.
	$ppPrintCurrent        = [Microsoft.Office.Interop.PowerPoint.PpPrintRangeType]::ppPrintCurrent        # Print the current slide from the presentation.
	$ppPrintSlideRange     = [Microsoft.Office.Interop.PowerPoint.PpPrintRangeType]::ppPrintSlideRange     # Print a range of slides.
	$ppPrintNamedSlideShow = [Microsoft.Office.Interop.PowerPoint.PpPrintRangeType]::ppPrintNamedSlideShow # Print a named slideshow.

	$ppShowAll            = [Microsoft.Office.Interop.PowerPoint.PpSlideShowRangeType]::ppShowAll             # Show all.
	$ppShowNamedSlideShow = [Microsoft.Office.Interop.PowerPoint.PpSlideShowRangeType]::ppShowNamedSlideShow  # Show named slideshow.
	$ppShowSlideRange     = [Microsoft.Office.Interop.PowerPoint.PpSlideShowRangeType]::ppShowSlideRange      # Show slide range.

    $PPTApp = New-Object "Microsoft.Office.Interop.Powerpoint.ApplicationClass"

}
Process
{
    foreach ($CurrentPath in $Path)
    {
        Get-ChildItem -Path $CurrentPath -Recurse -Filter *.ppt? | ForEach-Object {

        # Create a name for the PDF document; they are stored in the invocation folder!
        # If you want them to be created locally in the folders containing the source PowerPoint file, replace $curr_path with $_.DirectoryName
        $outputFile = [System.IO.Path]::ChangeExtension($_.FullName, ".pdf")
        if ($pscmdlet.ShouldProcess($_.FullName, "Convert PPT to PDF"))
        {
            # Open it in PowerPoint
            
            $presentation = $PPTApp.Presentations.Open($_.FullName, $msoTrue, $msoFalse, $msoFalse)
            #Bug in PPT, PrintRange is required even if optional..., cannot be null or 0, won't be used if RangeType set to PrintAll
            $ranges = $presentation.PrintOptions.Ranges
            $range = $ranges.Add(1,1)
            # Save it in PDF
            $presentation.ExportAsFixedFormat($outputFile,                      # File Name
                                              $ppFixedFormatTypePDF,            # File Type
                                              $ppFixedFormatIntentScreen,       # Screen or Print
                                              $msoTrue,                         # With or without frame
                                              $ppPrintHandoutHorizontalFirst,   # Vertical or Horizontal Handouts
                                              $ppPrintOutputSlides,             # What to export (slides, handouts etc.)
                                              $msoFalse,                        # Print Hidden slides or not
                                              $range,                           # Slide Range
                                              $ppPrintAll,                      # Range Type
                                              $null,                            # Slide show name
                                              $False,                           # Include document properties
                                              $False,                           # Keep IRM settings
                                              $False,                           # Keep document structure tags
                                              $True,                            # Include a bitmap of the text
                                              $False)                           # Use Iso 19005-1 (PDF/A)

            # Close PowerPoint file
            $presentation.Close()
            $presentation = $null
        }
    }
}

}
End
{
    # Exit and release the PowerPoint object
    $PPTApp.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($PPTApp)
    $PPTApp = $null
}
