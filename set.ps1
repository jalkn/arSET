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

    # Create templates directory structure
    $directories = @(
        "core/src",
        "core/static",
        "core/static/css",
        "core/static/js",
        "core/templates",
        "core/templates/admin",
        "core/templates/registration"
    )
    foreach ($dir in $directories) {
        New-Item -Path $dir -ItemType Directory -Force
    }

# Create models.py with cedula as primary key
Set-Content -Path "core/models.py" -Value @" 
"@

Set-Content -Path "core/views.py" -Value @"
from django.shortcuts import render, redirect
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.contrib.auth import get_user_model
from django.contrib.auth.mixins import LoginRequiredMixin
from django.views.generic import TemplateView

# The 'main' view is now a simple redirect to login if not authenticated, or a placeholder page if authenticated.
@login_required
def main(request):
    return render(request, 'home.html') # A simple home.html will be created below

def register_superuser(request):
    if request.method == 'POST':
        username = request.POST.get('username')
        email = request.POST.get('email')
        password1 = request.POST.get('password1')
        password2 = request.POST.get('password2')

        if password1 != password2:
            messages.error(request, "Passwords don't match")
            return redirect('register')

        User = get_user_model()
        if User.objects.filter(username=username).exists():
            messages.error(request, "Username already exists")
            return redirect('register')

        try:
            user = User.objects.create_superuser(
                username=username,
                email=email,
                password=password1
            )
            messages.success(request, f"Superuser {username} created successfully!")
            return redirect('login')
        except Exception as e:
            messages.error(request, f"Error creating superuser: {str(e)}")
            return redirect('register')

    return render(request, 'registration/register.html')

class ImportView(LoginRequiredMixin, TemplateView):
    template_name = 'import.html'
    
    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        # Add any additional context data you need
        return context
    
@login_required
def main(request):
    return render(request, 'home.html')
"@

# Create admin.py with enhanced configuration
Set-Content -Path "core/admin.py" -Value @" 
"@

# Create urls.py for core app
Set-Content -Path "core/urls.py" -Value @"
from django.contrib.auth import views as auth_views
from django.urls import path
from . import views
from django.contrib.auth import get_user_model
from django.contrib import messages
from django.shortcuts import render, redirect
from django.urls import path
from django.contrib.auth import views as auth_views
from .views import main, register_superuser, ImportView  

def register_superuser(request):
    if request.method == 'POST':
        username = request.POST.get('username')
        email = request.POST.get('email')
        password1 = request.POST.get('password1')
        password2 = request.POST.get('password2')
        
        if password1 != password2:
            messages.error(request, "Passwords don't match")
            return redirect('register')
        
        User = get_user_model()
        if User.objects.filter(username=username).exists():
            messages.error(request, "Username already exists")
            return redirect('register')
        
        try:
            user = User.objects.create_superuser(
                username=username,
                email=email,
                password=password1
            )
            messages.success(request, f"Superuser {username} created successfully!")
            return redirect('login')
        except Exception as e:
            messages.error(request, f"Error creating superuser: {str(e)}")
            return redirect('register')
    
    return render(request, 'registration/register.html')

urlpatterns = [
    path('', views.main, name='main'),
    path('logout/', auth_views.LogoutView.as_view(), name='logout'),
    path('register/', register_superuser, name='register'),
    path('import/', ImportView.as_view(), name='import'), 
]
"@

# Update project urls.py with proper admin configuration
Set-Content -Path "arpa/urls.py" -Value @"
from django.contrib import admin
from django.urls import include, path
from django.contrib.auth import views as auth_views

# Customize default admin interface
admin.site.site_header = 'A R P A'
admin.site.site_title = 'ARPA Admin Portal'
admin.site.index_title = 'Bienvenido a A R P A'

urlpatterns = [
    path('admin/', admin.site.urls),
    path('persons/', include('core.urls')),
    path('accounts/', include('django.contrib.auth.urls')),  
    path('', include('core.urls')), 
]
"@

#statics css style
@" 
:root {
    --primary-color: #0b00a2;
    --primary-hover: #090086;
    --text-on-primary: white;
}

body {
    margin: 0;
    padding: 20px;
    background-color: #f8f9fa;
}

.topnav-container {
    display: flex;
    align-items: center;
    padding: 0 40px;
    margin-bottom: 20px;
    gap: 15px;
}

.logoIN {
    width: 40px;
    height: 40px;
    background-color: var(--primary-color);
    border-radius: 8px;
    position: relative;
    flex-shrink: 0;
}

.logoIN::before {
    content: "";
    position: absolute;
    width: 100%;
    height: 100%;
    border-radius: 50%;
    top: 30%;
    left: 70%;
    transform: translate(-50%, -50%);
    background-image: linear-gradient(to right, 
        #ffffff 2px, transparent 2px);
    background-size: 4px 100%;
}

.navbar-title {
    color: var(--primary-color);
    font-weight: bold;
    font-size: 1.25rem;
    margin-right: auto;
}

.navbar-buttons {
    display: flex;
    gap: 10px;
}

.btn-custom-primary {
    background-color: white;
    border-color: var(--primary-color);
    color: var(--primary-color);
    padding: 0.5rem 1rem;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    min-width: 40px;
}

.btn-custom-primary:hover,
.btn-custom-primary:focus {
    background-color: var(--primary-hover);
    border-color: var(--primary-hover);
    color: var(--text-on-primary);
}

.btn-custom-primary i,
.btn-outline-dark i {
    margin-right: 0;
    font-size: 1rem;
    line-height: 1;
    display: inline-block;
    vertical-align: middle;
}

.main-container {
    padding: 0 40px;
}

/* Search filter styles */
.search-filter {
    margin-bottom: 20px;
    max-width: 400px;
}

/* Table row hover effect */
.table-hover tbody tr:hover {
    background-color: rgba(11, 0, 162, 0.05);
}

.btn-my-green {
    background-color: white;
    border-color: rgb(0, 166, 0);
    color: rgb(0, 166, 0);
}

.btn-my-green:hover {
    background-color: darkgreen;
    border-color: darkgreen;
    color: white;
}

.btn-my-green:focus,
.btn-my-green.focus {
    box-shadow: 0 0 0 0.2rem rgba(0, 128, 0, 0.5);
}

.btn-my-green:active,
.btn-my-green.active {
    background-color: darkgreen !important;
    border-color: darkgreen !important;
}

.btn-my-green:disabled,
.btn-my-green.disabled {
    background-color: lightgreen;
    border-color: lightgreen;
    color: #6c757d;
}

/* Card styles */
.card {
    border-radius: 8px;
    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
}

/* Table styles */
.table {
    width: 100%;
    margin-bottom: 1rem;
    color: #212529;
}

.table th {
    vertical-align: bottom;
    border-bottom: 2px solid #dee2e6;
}

.table td {
    vertical-align: middle;
}

/* Alert styles */
.alert {
    position: relative;
    padding: 0.75rem 1.25rem;
    margin-bottom: 1rem;
    border: 1px solid transparent;
    border-radius: 0.25rem;
}

/* Badge styles */
.badge {
    display: inline-block;
    padding: 0.35em 0.65em;
    font-size: 0.75em;
    font-weight: 700;
    line-height: 1;
    color: #fff;
    text-align: center;
    white-space: nowrap;
    vertical-align: baseline;
    border-radius: 0.25rem;
}

.bg-success {
    background-color:rgb(0, 166, 0) !important;
}

.bg-danger {
    background-color: #dc3545 !important;
}

"@ | Out-File -FilePath "core/static/css/style.css" -Encoding utf8

@"
.loading-overlay {
    position: fixed;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    background-color: rgba(0,0,0,0.5);
    z-index: 9999;
    display: none;
    justify-content: center;
    align-items: center;
}

.loading-content {
    background-color: white;
    padding: 30px;
    border-radius: 8px;
    text-align: center;
    max-width: 500px;
    width: 90%;
}

.progress {
    height: 20px;
    margin: 20px 0;
}

/* Spinner styles for submit buttons */
.btn .spinner-border {
    margin-right: 8px;
}
"@ | Out-File -FilePath "core/static/css/loading.css" -Encoding utf8

# Create loading.js
@"
document.addEventListener('DOMContentLoaded', function() {
    // Get all forms that should show loading
    const forms = document.querySelectorAll('form');
    
    forms.forEach(form => {
        form.addEventListener('submit', function(e) {
            // Only show loading for forms that aren't the search form
            if (!form.classList.contains('no-loading')) {
                // Show loading overlay
                const loadingOverlay = document.getElementById('loadingOverlay');
                if (loadingOverlay) {
                    loadingOverlay.style.display = 'flex';
                }
                
                // Optional: Disable submit button to prevent double submission
                const submitButton = form.querySelector('button[type="submit"]');
                if (submitButton) {
                    submitButton.disabled = true;
                    submitButton.innerHTML = '<span class="spinner-border spinner-border-sm" role="status" aria-hidden="true"></span> Procesando...';
                }
            }
        });
    });
});
"@ | Out-File -FilePath "core/static/js/loading.js" -Encoding utf8

# Create custom admin base template
@"
{% extends "admin/base.html" %}

{% block title %}{{ title }} | {{ site_title|default:_('A R P A') }}{% endblock %}

{% block branding %}
<h1 id="site-name"><a href="{% url 'admin:index' %}">{{ site_header|default:_('A R P A') }}</a></h1>
{% endblock %}

{% block nav-global %}{% endblock %}
"@ | Out-File -FilePath "core/templates/admin/base_site.html" -Encoding utf8

# Create master template
@"
<!DOCTYPE html>
<html>
<head>
    <title>{% block title %}ARPA{% endblock %}</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    {% load static %}
    <link rel="stylesheet" href="{% static 'css/style.css' %}">
</head>
<body>
    {% if user.is_authenticated %}
    <div class="topnav-container">
        <a href="/" style="text-decoration: none;">
            <div class="logoIN"></div>
        </a>
        <div class="navbar-title">{% block navbar_title %}ARPA{% endblock %}</div>
        <div class="navbar-buttons">
            {% block navbar_buttons %}
            <a href="/admin/" class="btn btn-outline-dark" title="Admin">
                <i class="fas fa-wrench"></i>
            </a>
            <a href="{% url 'import' %}" class="btn btn-custom-primary" title="Importar">
                <i class="fas fa-database"></i>
            </a>
            <form method="post" action="{% url 'logout' %}" class="d-inline">
                {% csrf_token %}
                <button type="submit" class="btn btn-custom-primary" title="Cerrar sesiÃ³n">
                    <i class="fas fa-sign-out-alt"></i>
                </button>
            </form>
            {% endblock %}
        </div>
    </div>
    {% endif %}
    
    <div class="main-container">
        {% if messages %}
            {% for message in messages %}
                <div class="alert alert-{{ message.tags }} alert-dismissible fade show">
                    {{ message }}
                    <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
                </div>
            {% endfor %}
        {% endif %}
        
        {% block content %}
        {% endblock %}
    </div>
    
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>
"@ | Out-File -FilePath "core/templates/master.html" -Encoding utf8

# Create home template
@"
{% extends "master.html" %}

{% block title %}A R P A{% endblock %}
{% block navbar_title %}Dashboard{% endblock %}

{% block navbar_buttons %}
<div>
    <a href="/persons/" class="btn btn-custom-primary">
        <i class="fas fa-users"></i>
    </a>
    <a href="/finance/" class="btn btn-custom-primary">
        <i class="fas fa-chart-line" style="color: green;"></i>
    </a>
    <a href="/cards/" class="btn btn-custom-primary" title="Tarjetas">
        <i class="far fa-credit-card" style="color: blue;"></i>
    </a>
    <a href="/conflicts/" class="btn btn-custom-primary">
        <i class="fas fa-balance-scale" style="color: orange;"></i>
    </a>
    <a href="/alerts/" class="btn btn-custom-primary">
        <i class="fas fa-bell" style="color: red;"></i>
    </a>
    <form method="post" action="{% url 'logout' %}" class="d-inline">
        {% csrf_token %}
        <button type="submit" class="btn btn-custom-primary" title="Cerrar sesiÃƒÂ³n">
            <i class="fas fa-sign-out-alt"></i>
        </button>
    </form>
</div>
{% endblock %}

{% block content %}
<div class="row justify-content-center">
    <div class="col-md-6">
        <div class="card border-0 shadow">
            <div class="card-body p-5">
                <div style="align-items: center; text-align: center;"> 
                    <h5 class="card-title">Bienvenido a ARPA</h5>
                    <p class="card-text">Automatizacion Robotica de Procesos de Auditoria</p>
                    <a href="{% url 'import' %}" class="btn btn-custom-primary">
                        <i class="fas fa-upload"></i> 
                    </a>
                </div>
            </div>
        </div>
    </div>
</div>
{% endblock %}
"@ | Out-File -FilePath "core/templates/home.html" -Encoding utf8

# Create login template
@"
{% extends "master.html" %}

{% block title %}ARPA{% endblock %}
{% block navbar_title %}ARPA{% endblock %}

{% block navbar_buttons %}
{% endblock %}

{% block content %}
<div class="row justify-content-center">
    <div class="col-md-6">
        <div class="card border-0 shadow">
            <div class="card-body p-5">
                <div style="align-items: center; text-align: center;"> 
                        <a href="/" style="text-decoration: none;" >
                            <div class="logoIN" style="margin: 20px auto;"></div>
                        </a>
                    {% if form.errors %}
                    <div class="alert alert-danger">
                        Tu nombre de usuario y clave no coinciden. Por favor intenta de nuevo.
                    </div>
                    {% endif %}

                    {% if next %}
                        {% if user.is_authenticated %}
                        <div class="alert alert-warning">
                            Tu cuenta no tiene acceso a esta pagina. Para continuar,
                            por favor ingresa con una cuenta que tenga acceso.
                        </div>
                        {% else %}
                        <div class="alert alert-info">
                            Por favor accede con tu clave para ver esta pagina.
                        </div>
                        {% endif %}
                    {% endif %}

                    <form method="post" action="{% url 'login' %}">
                        {% csrf_token %}

                        <div class="mb-3">
                            <input type="text" name="username" class="form-control form-control-lg" id="id_username" placeholder="Usuario" required>
                        </div>

                        <div class="mb-4">
                            <input type="password" name="password" class="form-control form-control-lg" id="id_password" placeholder="Clave" required>
                        </div>

                        <div class="d-flex align-items-center justify-content-between">
                            <button type="submit" class="btn btn-custom-primary btn-lg">
                                <i class="fas fa-sign-in-alt"style="color: green;"></i>
                            </button>
                            <div>
                                <a href="{% url 'register' %}" class="btn btn-custom-primary" title="Registrarse">  
                                    <i class="fas fa-user-plus fa-lg"></i>
                                </a>
                                <a href="{% url 'password_reset' %}" class="btn btn-custom-primary" title="Recupera tu acceso">
                                    <i class="fas fa-key fa-lg" style="color: orange;"></i>
                                </a>
                            </div>
                        </div>

                        <input type="hidden" name="next" value="{{ next }}">
                    </form>
                </div> 
            </div>
        </div>
    </div>
</div>
{% endblock %}
"@ | Out-File -FilePath "core/templates/registration/login.html" -Encoding utf8

# Create register template
@"
{% extends "master.html" %}

{% block title %}Registro{% endblock %}
{% block navbar_title %}Registro{% endblock %}

{% block navbar_buttons %}
{% endblock %}

{% block content %}
<div class="row justify-content-center">
    <div class="col-md-6">
        <div class="card border-0 shadow">
            <div class="card-body p-5">
                <div style="align-items: center; text-align: center;"> 
                        <a href="/" style="text-decoration: none;" >
                            <div class="logoIN" style="margin: 20px auto;"></div>
                        </a>
                    {% if messages %}
                        {% for message in messages %}
                            <div class="alert alert-{% if message.tags == 'error' %}danger{% else %}{{ message.tags }}{% endif %}">
                                {{ message }}
                            </div>
                        {% endfor %}
                    {% endif %}

                    <form method="post" action="{% url 'register' %}">
                        {% csrf_token %}
                        <div class="mb-3">
                            <input type="text" name="username" class="form-control form-control-lg" id="username" placeholder="Usuario" required>
                        </div>

                        <div class="mb-3">
                            <input type="email" name="email" class="form-control form-control-lg" id="email" placeholder="Correo" required>
                        </div>

                        <div class="mb-3">
                            <input type="password" name="password1" class="form-control form-control-lg" id="password1" placeholder="Clave" required>
                        </div>

                        <div class="mb-3">
                            <input type="password" name="password2" class="form-control form-control-lg" id="password2" placeholder="Repite tu clave" required>
                        </div>

                        <div class="d-flex align-items-center justify-content-between">
                            <button type="submit" class="btn btn-custom-primary btn-lg">
                                <i class="fas fa-user-plus fa-lg" style="color: green;"></i>
                            </button>
                            <div>
                                <a href="{% url 'login' %}" class="btn btn-custom-primary" title="Ingresar">
                                    <i class="fas fa-sign-in-alt" style="color: rgb(0, 0, 255);"></i>
                                </a>
                                <a href="{% url 'password_reset' %}" class="btn btn-custom-primary" title="Recupera tu acceso">
                                    <i class="fas fa-key fa-lg" style="color: orange;"></i>
                                </a>
                            </div>
                        </div>

                        <input type="hidden" name="next" value="{{ next }}">
                    </form>
                </div> 
            </div>
        </div>
    </div>
</div>
{% endblock %}
"@ | Out-File -FilePath "core/templates/registration/register.html" -Encoding utf8

# Create import template
@"
{% extends "master.html" %}

{% block title %}Importar desde Excel{% endblock %}
{% block navbar_title %}Importar Datos{% endblock %}

{% block navbar_buttons %}
<div>
    <a href="/persons/" class="btn btn-custom-primary">
        <i class="fas fa-users"></i>
    </a>
    <a href="/finance/" class="btn btn-custom-primary">
        <i class="fas fa-chart-line" style="color: green;"></i>
    </a>
    <a href="/cards/" class="btn btn-custom-primary" title="Tarjetas">
        <i class="far fa-credit-card" style="color: blue;"></i>
    </a>
    <a href="/conflicts/" class="btn btn-custom-primary">
        <i class="fas fa-balance-scale" style="color: orange;"></i>
    </a>
    <a href="/alerts/" class="btn btn-custom-primary">
        <i class="fas fa-bell" style="color: red;"></i>
    </a>
    <form method="post" action="{% url 'logout' %}" class="d-inline">
        {% csrf_token %}
        <button type="submit" class="btn btn-custom-primary" title="Cerrar sesiÃ³n">
            <i class="fas fa-sign-out-alt"></i>
        </button>
    </form>
</div>
{% endblock %}

{% block content %}
{% load static %}
<div class="loading-overlay" id="loadingOverlay">
    <div class="loading-content">
        <h4>Procesando datos...</h4>
        <div class="progress">
            <div class="progress-bar progress-bar-striped progress-bar-animated" 
                 role="progressbar" 
                 style="width: 100%"></div>
        </div>
        <p>Por favor espere, esto puede tomar unos segundos.</p>
    </div>
</div>

<!-- Add loading CSS -->
<link rel="stylesheet" href="{% static 'css/loading.css' %}">

<!-- Add loading JS -->
<script src="{% static 'js/loading.js' %}"></script>

<div class="row">
    <!-- First row with 3 cards -->
    <div class="col-md-4 mb-4">
        <div class="card h-100">
            <div class="card-body">
                <form method="post" enctype="multipart/form-data">
                    {% csrf_token %}
                    <div class="mb-3">
                        <input type="file" class="form-control" id="period_excel_file" name="period_excel_file" required>
                        <div class="form-text">El archivo Excel de Periodos debe incluir las columnas: Id, Activo, FechaFinDeclaracion, FechaInicioDeclaracion, Ano declaracion</div>
                    </div>
                    <button type="submit" class="btn btn-custom-primary btn-lg text-start">Importar Periodos</button>
                </form>
            </div>
            {% for message in messages %}
                {% if 'import_period_excel' in message.tags %}
                <div class="card-footer">
                    <div class="alert alert-{{ message.tags }} alert-dismissible fade show mb-0">      
                        {{ message }}
                        <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
                    </div>
                </div>
                {% endif %}
            {% endfor %}
        </div>
    </div>

    <div class="col-md-4 mb-4">
        <div class="card h-100">
            <div class="card-body">
                <form method="post" enctype="multipart/form-data"> 
                    {% csrf_token %}
                    <div class="mb-3">
                        <input type="file" class="form-control" id="excel_file" name="excel_file" required>
                        <div class="form-text">El archivo Excel de Personas debe incluir las columnas: Id, NOMBRE COMPLETO, CARGO, Cedula, Correo, Compania, Estado</div>
                    </div>
                    <button type="submit" class="btn btn-custom-primary btn-lg text-start">Importar Personas</button>
                </form>
            </div>
            {% for message in messages %}
                {% if 'import_persons' in message.tags %}
                <div class="card-footer">
                    <div class="alert alert-{{ message.tags }} alert-dismissible fade show mb-0">      
                        {{ message }}
                        <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
                    </div>
                </div>
                {% endif %}
            {% endfor %}
        </div>
    </div>

    <div class="col-md-4 mb-4">
        <div class="card h-100">
            <div class="card-body">
                <form method="post" enctype="multipart/form-data">
                    {% csrf_token %}
                    <div class="mb-3">
                        <input type="file" class="form-control" id="conflict_excel_file" name="conflict_excel_file" required>
                        <div class="form-text">'ID', 'Cedula', 'Nombre', 'Compania', 'Cargo', 'Email', 'Fecha de Inicio', 'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Q9', 'Q10', 'Q11' </div>
                    </div>
                    <button type="submit" class="btn btn-custom-primary btn-lg text-start">Importar Conflictos</button>
                </form>
            </div>
            {% for message in messages %}
                {% if 'import_conflict_excel' in message.tags %}
                <div class="card-footer">
                    <div class="alert alert-{{ message.tags }} alert-dismissible fade show mb-0">      
                        {{ message }}
                        <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
                    </div>
                </div>
                {% endif %}
            {% endfor %}
        </div>
    </div>
</div>

<div class="row">
    <!-- Left column with 3 import forms in a single card -->
    <div class="col-md-4">
        <div class="card h-100">
            <div class="card-body d-flex flex-column">
                <!-- Bienes y Rentas -->
                <div class="mb-4 flex-grow-1">
                    <form method="post" enctype="multipart/form-data">
                        {% csrf_token %}
                        <div class="mb-3">
                            <input type="file" class="form-control" id="protected_excel_file" name="protected_excel_file" required>
                            <div class="form-text">El archivo Excel de Bienes y Rentas debe incluir las columnas: </div>
                            <div class="mb-3">
                                <input type="password" class="form-control" id="excel_password" name="excel_password">
                                <div class="form-text">Ingrese la contrasena</div>
                            </div>
                        </div>
                        <button type="submit" class="btn btn-custom-primary btn-lg text-start">Importar Bienes y Rentas</button>
                    </form>
                </div>

                <!-- Visa -->
                <div class="flex-grow-1">
                    <form method="post" enctype="multipart/form-data">
                        {% csrf_token %}
                        <div class="mb-3">
                            <input type="file" class="form-control" id="visa_pdf_files" name="visa_pdf_files" multiple webkitdirectory directory required>
                            <div class="form-text">Seleccione la carpeta con los PDFs de VISA</div>
                        </div>
                        <!-- Add password input field -->
                        <div class="mb-3">
                            <input type="password" class="form-control" id="visa_pdf_password" name="visa_pdf_password" placeholder="Clave">
                            <div class="form-text">Ingrese la contraseña si los PDFs están protegidos</div>
                        </div>
                        <button type="submit" class="btn btn-custom-primary btn-lg text-start">Procesar VISA</button>
                    </form>
                </div>
            </div>
            
            <!-- Messages for all three forms -->
            {% for message in messages %}
                {% if 'import_protected_excel' in message.tags or 'import_visa_pdfs' in message.tags %}
                <div class="card-footer">
                    <div class="alert alert-{{ message.tags }} alert-dismissible fade show mb-0">
                        {{ message }}
                        <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
                    </div>
                </div>
                {% endif %}
            {% endfor %}
        </div>
    </div>

    <!-- Right column with analysis results -->
    <div class="col-md-8">
        <div class="card h-100">
            <div class="card-header bg-light">
                <h5 class="mb-0">Resultados del Analisis</h5>
            </div>
            <div class="card-body">
                {% if analysis_results %}
                <div class="table-responsive">
                    <table class="table table-sm">
                        <thead>
                            <tr>
                                <th>Archivo Generado</th>
                                <th>Registros</th>
                                <th>Estado</th>
                                <th>Ultima Actualizacion</th>
                            </tr>
                        </thead>
                        <tbody>
                            {% for result in analysis_results %}
                            <tr>
                                <td>{{ result.filename }}</td>
                                <td>{{ result.records|default:"-" }}</td>
                                <td>
                                    <span class="badge bg-{% if result.status == 'success' %}success{% elif result.status == 'error' %}danger{% else %}secondary{% endif %}">
                                        {% if result.status == 'success' %}
                                            Exitoso
                                        {% elif result.status == 'pending' %}
                                            Pendiente
                                        {% elif result.status == 'error' %}
                                            Error
                                        {% else %}
                                            {{ result.status|capfirst }}
                                        {% endif %}
                                    </span>
                                    {% if result.status == 'error' and result.error %}
                                    <small class="text-muted d-block">{{ result.error }}</small>   
                                    {% endif %}
                                </td>
                                <td>
                                    {% if result.last_updated %}
                                    {{ result.last_updated|date:"d/m/Y H:i" }}
                                    {% else %}
                                    -
                                    {% endif %}
                                </td>
                            </tr>
                            {% endfor %}
                        </tbody>
                    </table>
                </div>
                {% else %}
                <div class="text-center py-4">
                    <i class="fas fa-info-circle fa-3x text-muted mb-3"></i>
                    <p class="text-muted">No hay resultados de analisis disponibles</p>
                </div>
                {% endif %}
            </div>
            <div class="card-footer">
                <small class="text-muted">Los archivos se procesan en: core/src/</small>
            </div>
        </div>
    </div>
</div>
{% endblock %}
"@ | Out-File -FilePath "core/templates/import.html" -Encoding utf8


    # Update settings.py
    $settingsContent = Get-Content -Path ".\arpa\settings.py" -Raw
    $settingsContent = $settingsContent -replace "INSTALLED_APPS = \[", "INSTALLED_APPS = [
    'core.apps.CoreConfig',
    'django.contrib.humanize',"
    $settingsContent = $settingsContent -replace "from pathlib import Path", "from pathlib import Path
import os"
    $settingsContent | Set-Content -Path ".\arpa\settings.py"

# Add static files configuration
Add-Content -Path ".\arpa\settings.py" -Value @"

# Static files (CSS, JavaScript, Images)
STATIC_URL = 'static/'
STATIC_ROOT = BASE_DIR / 'staticfiles'
STATICFILES_DIRS = [
    BASE_DIR / "core/static",
]

MEDIA_URL = 'media/'
MEDIA_ROOT = BASE_DIR / 'media'

# Custom admin skin
ADMIN_SITE_HEADER = "A R P A"
ADMIN_SITE_TITLE = "ARPA Admin Portal"
ADMIN_INDEX_TITLE = "Bienvenido a A R P A"

LOGIN_REDIRECT_URL = '/'  # Where to redirect after login
LOGOUT_REDIRECT_URL = '/accounts/login/'  # Where to redirect after logout
"@

    # Run migrations
    python manage.py makemigrations core
    python manage.py migrate

    # Create superuser
    #python manage.py createsuperuser

    python manage.py collectstatic --noinput

    # Start the server
    Write-Host "🚀 Starting Django development server..." -ForegroundColor $GREEN
    python manage.py runserver

}

arpa