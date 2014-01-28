rem enter Temp Folder of PDF-Mailer
set PDFMailerTemp="D:\PDFMailer\Temp\"

rem remove temporary PDF-Mailer files
rmdir /s /q %PDFMailerTemp%

rem recreate TEMP-folder
mkdir %PDFMailerTemp%

