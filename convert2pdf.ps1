


# function GetPdffile{
#     param(
#         $file,
#         $pdf_folder
#     )
    
#     $path = Get-Item -path $file
    
#     if (-not $pdf_folder){
#         write-output $pdf_folder 'not found'
#         $pdf_file = Join-Path -Path $path.DirectoryName -ChildPath $($path.name.split(".")[0] + ".pdf")
#         return $pdf_file
#     }
#     else{
#         $pdf_path = Join-Path -Path $path.DirectoryName -ChildPath $pdf_folder 
#         $pdf_file = Join-Path -path $pdf_path -ChildPath $($path.name.split(".")[0] + ".pdf")
#         write-output $pdf_file
#         return $pdf_file
#     }
# }

Function WordConvertToPDF {
    param(
        $wordfile,
        $pdf_folder
    )
    # single file
    $file = Get-Item -path $wordfile
    Write-Output "processing: $($file.Name)"
    $word_app = New-Object -ComObject Word.Application
    $word_app.visible = $false
    $document = $word_app.Documents.Open($file.FullName)
    # $pdf_file = GetPdffile $file $pdf_folder
    $pdf_file = "$($file.DirectoryName)\\$pdf_folder\\$($file.BaseName).pdf"
    # write-output $pdf_file
    $document.SaveAs([ref] $pdf_file, [ref] 17)

    $document.Close()
    $word_app.Quit()
    [gc]::Collect();
    [gc]::WaitForPendingFinalizers();
    Write-Output "done"
}

Function PptConvertToPDF {
    param(
        $pptfile,
        $pdf_folder
    )
    # single file
    $ppt_app = New-Object -ComObject PowerPoint.Application
    # $ppt_app.visible = [Microsoft.Office.Core.MsoTriState]::msoFalse

    $file = Get-Item -path $pptfile
    Write-Output "processing: $($file.Name)"
    $ppt = $ppt_app.Presentations.Open($file.FullName)
    # $pdf_file = GetPdffile $file $pdf_folder 
    $pdf_file = "$($file.DirectoryName)\\$pdf_folder\\$($file.BaseName).pdf"
    # write-output $pdf_file
    $ppt.SaveAs([ref] $pdf_file, [ref] 32)

    $ppt_app.Quit()
    $ppt_app = $null
    [gc]::Collect();
    [gc]::WaitForPendingFinalizers();
    Write-Output "done"
}

Function ScanWord2PDF {
    param(
        $folder,
        $pdf_folder
    )
    $word_app = New-Object -ComObject Word.Application
    $word_app.visible = $false
    # This filter will find .doc as well as .docx documents
    Get-ChildItem -Path $folder -Filter *.doc* | ForEach-Object {

        Write-Output "processing: $_"

        $document = $word_app.Documents.Open($_.FullName)

        $pdf_filename = "$($_.DirectoryName)\\$pdf_folder\\$($_.BaseName).pdf"

        $document.SaveAs([ref] $pdf_filename, [ref] 17)

        $document.Close()
    }

    $word_app.Quit()
    $word_app = $null
    [gc]::Collect();
    [gc]::WaitForPendingFinalizers();

}

Function ScanPPT2PDF {
    param(
        $folder,
        $pdf_folder
    )
    $ppt_app = New-Object -ComObject PowerPoint.Application
    # $ppt_app.visible = [Microsoft.Office.Core.MsoTriState]::msoFalse
    # This filter will find .ppt as well as .pptx documents
    Get-ChildItem -Path $folder -Filter *.ppt* | ForEach-Object {

        Write-Output "processing: $_"

        $presentation = $ppt_app.Presentations.Open($_.FullName)

        $pdf_filename = "$($_.DirectoryName)\\$pdf_folder\\$($_.BaseName).pdf"

        $presentation.SaveAs([ref] $pdf_filename, [ref] 32)

        $presentation.Close()
    }
    $ppt_app.Quit()
    $ppt_app = $null
    [gc]::Collect();
    [gc]::WaitForPendingFinalizers();
}

# Add the PowerPoint/Word assemblies that we'll need
Add-type -AssemblyName office -ErrorAction SilentlyContinue
Add-Type -AssemblyName microsoft.office.interop.powerpoint -ErrorAction SilentlyContinue
Add-Type -AssemblyName microsoft.office.interop.word -ErrorAction SilentlyContinue

$pdf_sub_folder_name = "converted_pdfs"

# main function entry
function Main {


    param($file = $pwd)

    # Write-Host "The selected file is: $file"
    $path = Get-Item -path $file

    if ($path.PSIsContainer) {
        # 这是一个目录
        # Write-Host "Directory:" $path
        # 在这里添加对目录的操作
        New-Item -Path $path\$pdf_sub_folder_name -ItemType directory -Force
        ScanPPT2PDF $path $pdf_sub_folder_name
        ScanWord2PDF $path $pdf_sub_folder_name
    }
    else {
        # 这是一个文件
   
        # Write-Host "File:" $path.DirectoryName $path.Name
        New-Item -Path $(Join-Path -path $path.DirectoryName -childpath $pdf_sub_folder_name) -ItemType directory -Force
        if ($path.Extension -eq ".doc" -or $path.Extension -eq ".docx") {
            WordConvertToPDF $path $pdf_sub_folder_name
        }
        elseif ($path.Extension -eq ".ppt" -or $path.Extension -eq ".pptx") {
            PptConvertToPDF $path $pdf_sub_folder_name
        }
        else {
            Write-Host "Not a word or ppt file"
        }

    }

    Read-Host -Prompt "Press any key to continue..." | Out-Null
}
main $args[0]