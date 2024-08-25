@echo off
setlocal enabledelayedexpansion

:: Define output file
set "finalFile=final_file.csv"

:: Create/clear the final file
type nul > "!finalFile!"

:: Variable to handle the header
set "first=1"

:: Loop to process each CSV file in the directory
for %%f in (*.csv) do (
  if "%%f" neq "!finalFile!" (
    if !first! equ 1 (
      :: Copy the first file completely, including the headers
      type "%%f" >> "!finalFile!"
      set "first=0"
    ) else (
      :: Add the following files excluding the first line
      more +1 "%%f" >> "!finalFile!"
    )
  )
)

:: Call PowerShell to sort the final file by TimeInt column
powershell -Command "& { $file = Import-Csv -Path '.\!finalFile!'; $sorted = $file | Sort-Object -Property TimeInt; $sorted | Export-Csv -Path '.\!finalFile!' -NoTypeInformation -Force; }"

echo Todos los archivos CSV han sido unidos y ordenados en !finalFile!

endlocal
pause