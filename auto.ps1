if (-not $env:VIRTUAL_ENV) {
    & ".\venv\Scripts\Activate.ps1"
    Write-Host "Ambiente virtual ativado."
} else {
    Write-Host "Ambiente já está ativado."
}

python .\file_organizer_v2.py