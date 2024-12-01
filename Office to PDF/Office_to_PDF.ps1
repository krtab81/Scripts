# Get the command line arguments for the script
param (
    [Parameter(Mandatory = $false)]
    [string[]]$FilePaths
)

# -------------------
# CONFIGURATION START
# -------------------

# Display a summary when the conversions are complete
$SUMMARY_DISPLAY = $true
$SUMMARY_TITLE = "Conversion Complete"

# File extensio for PDFs
$PDF_Extension = "pdf"

# Results for CheckFile function
$CHECKFILE_OK = 0
$CHECKFILE_FileDoesNotExist = 1
$CHECKFILE_NotMSOFile = 2

# Settings to produce PDFs from the applications
$EXCEL_PDF = 0
$EXCEL_QualityStandard = 0
$WORD_PDF = 17
$POWERPOINT_PDF = 32

# File types returned from OfficeFileType function
$FILE_TYPE_NotOffice = 0
$FILE_TYPE_Word = 1
$FILE_TYPE_Excel = 2
$FILE_TYPE_PowerPoint = 3

# Valid file type lists
$g_strFileTypesWord = @("DOC", "DOCX")
$g_strFileTypesExcel = @("XLS", "XLSX")
$g_strFileTypesPowerPoint = @("PPT", "PPTX")

	
# --------------------
# CONFIGURATION FINISH
# --------------------


function CheckFile ($filePath) {
    # Check file exists and is an office file (Excel, Word, PowerPoint)
    if (IsOfficeFile $filePath) {
        if (Test-Path $filePath) {
            return $CHECKFILE_OK
        } else {
            return $CHECKFILE_FileDoesNotExist
        }
    } else {
        return $CHECKFILE_NotMSOFile
    }
}

function OfficeFileType ($filePath) {
    # 'Returns the type of office file, based upon file extension
    $extension = (Get-Item $filePath).Extension.TrimStart(".").ToUpper()

    if ($g_strFileTypesWord -contains $extension) {
        return $FILE_TYPE_Word
    } elseif ($g_strFileTypesExcel -contains $extension) {
        return $FILE_TYPE_Excel
    } elseif ($g_strFileTypesPowerPoint -contains $extension) {
        return $FILE_TYPE_PowerPoint
    } else {
        return $FILE_TYPE_NotOffice
    }
}

function IsOfficeFile ($filePath) {
    # Returns true if a file is an office file (Excel, Word, PowerPoint)
    return (OfficeFileType $filePath -ne $FILE_TYPE_NotOffice)
}

function SaveWordAsPDF ($filePath) {
    # Returns true if a file is a Word file
    $word = New-Object -ComObject Word.Application
    $doc = $word.Documents.Open($filePath)
    $pdfPath = PathOfPDF $filePath
    #$doc.SaveAs($pdfPath, [ref]$WORD_PDF)	
    #$doc.SaveAs2($pdfPath, [ref]$WORD_PDF)
    $doc.ExportAsFixedFormat($WORD_PDF, $pdfPath)
    #$word.ActiveDocument.ExportAsFixedFormat($WORD_PDF, $pdfPath, $EXCEL_QualityStandard, $true, $false)
    $doc.Close()
    $word.Quit()
}

function SaveExcelAsPDF ($filePath) {
    # Returns true if a file is an Excel file
    $excel = New-Object -ComObject Excel.Application
    $workbook = $excel.Workbooks.Open($filePath)
    $pdfPath = PathOfPDF $filePath
    $workbook.ExportAsFixedFormat($EXCEL_PDF, $pdfPath)
    #, $EXCEL_QualityStandard, $true, $false)
    $workbook.Close($false)
    $excel.Quit()
}

function SavePowerPointAsPDF ($filePath) {
    # Returns true if a file is a PowerPoint file
    $powerpoint = New-Object -ComObject PowerPoint.Application
    $presentation = $powerpoint.Presentations.Open($filePath, $false, $false, $false)
    $pdfPath = PathOfPDF $filePath
    $presentation.SaveAs($pdfPath, [ref]$POWERPOINT_PDF)
    $presentation.Close()
    $powerpoint.Quit()
}

function PathOfPDF ($originalFilePath) {
    # Search a delimited list for a text string and return true if it's found
    $directory = Split-Path $originalFilePath
    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($originalFilePath)
    return Join-Path $directory "$fileName.$PDF_Extension"
}


# MAIN

if ($FilePaths.Count -eq 0) {
    # Write-Host "Por favor, proporcione al menos un archivo para procesar." -ForegroundColor Yellow
    Write-Host "Please pass a file to this script." -ForegroundColor Yellow		
    exit
}

foreach ($filePath in $FilePaths) {
    # Check we have a valid file and process it
    switch (CheckFile $filePath) {
        $CHECKFILE_OK {
            switch (OfficeFileType $filePath) {
                $FILE_TYPE_Word { SaveWordAsPDF $filePath }
                $FILE_TYPE_Excel { SaveExcelAsPDF $filePath }
                $FILE_TYPE_PowerPoint { SavePowerPointAsPDF $filePath }
            }
        }
        $CHECKFILE_FileDoesNotExist {
            # Write-Host "El archivo '$filePath' no existe." -ForegroundColor Red
            Write-Host "'$filePath' does not exist." -ForegroundColor Red			
            exit
        }
        $CHECKFILE_NotMSOFile {
            # Write-Host "El archivo '$filePath' no es un tipo de archivo válido." -ForegroundColor Red
            Write-Host "'$filePath' is not a valid file type." -ForegroundColor Red			
            exit
        }
    }
}

# Display an optional summary message
if ($SUMMARY_DISPLAY) {
    $fileCount = $FilePaths.Count
    # Write-Host "$fileCount archivo(s) convertido(s)." -ForegroundColor Green
    Write-Host "$fileCount file(s) converted(s)." -ForegroundColor Green	
}
