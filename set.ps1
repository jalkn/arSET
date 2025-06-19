function arpa {
    param (
        [string]$ExcelFilePath = $null
    )

    $YELLOW = [ConsoleColor]::Yellow
    $GREEN = [ConsoleColor]::Green

    Write-Host "🚀 Creating ARPA" -ForegroundColor $YELLOW

    # Create Python virtual environment
    python -m venv .venv
    .\.venv\scripts\activate

    # Install required Python packages
    python.exe -m pip install --upgrade pip
    python -m pip install django whitenoise django-bootstrap-v5 openpyxl pandas xlrd>=2.0.1 pdfplumber fitz

    # Create Django project
    django-admin startproject arpa
    cd arpa

    # Create core app
    python manage.py startapp core

}

arpa