# Definir las rutas de los archivos de origen y destino
$documentoOrigen = "C:\Users\User\Documents\crack.docx"
$documentoDestino = "C:\Users\User\Documents\cracked.docx"

# Crear una instancia de la aplicación Word
$wordApp = New-Object -ComObject Word.Application
$wordApp.Visible = $false

# Abrir el documento de origen y copiar su contenido
$documentoOrigenObj = $wordApp.Documents.Open($documentoOrigen)
$contenido = $documentoOrigenObj.Content.Text
$documentoOrigenObj.Close()

# Crear un nuevo documento de destino o abrir el existente
if (Test-Path $documentoDestino) {
    $documentoDestinoObj = $wordApp.Documents.Open($documentoDestino)
} else {
    $documentoDestinoObj = $wordApp.Documents.Add()
}

# Pegar el contenido en el documento de destino
$documentoDestinoObj.Content.Text = $contenido

# Guardar y cerrar el documento de destino
$documentoDestinoObj.SaveAs([ref] $documentoDestino)
$documentoDestinoObj.Close()

# Cerrar la aplicación Word
$wordApp.Quit()

# Liberar los objetos COM
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($documentoOrigenObj) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($documentoDestinoObj) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordApp) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Output "El contenido del documento ha sido copiado exitosamente."
