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
from django.db import models

class Person(models.Model):
    ESTADO_CHOICES = [
        ('Activo', 'Activo'),
        ('Retirado', 'Retirado'),
    ]
    
    cedula = models.CharField(max_length=20, primary_key=True, verbose_name="Cedula")
    nombre_completo = models.CharField(max_length=255, verbose_name="Nombre Completo")
    cargo = models.CharField(max_length=255, verbose_name="Cargo")
    correo = models.EmailField(max_length=255, verbose_name="Correo")
    compania = models.CharField(max_length=255, verbose_name="Compania")
    estado = models.CharField(max_length=20, choices=ESTADO_CHOICES, default='Activo', verbose_name="Estado")
    revisar = models.BooleanField(default=False, verbose_name="Revisar")
    comments = models.TextField(blank=True, null=True, verbose_name="Comentarios")

    def __str__(self):
        return f"{self.cedula} - {self.nombre_completo}"

    class Meta:
        verbose_name = "Persona"
        verbose_name_plural = "Personas"

class FinancialReport(models.Model):
    person = models.ForeignKey(Person, on_delete=models.CASCADE, related_name='financial_reports')
    fkIdPeriodo = models.CharField(max_length=20, blank=True, null=True)
    ano_declaracion = models.CharField(max_length=20, blank=True, null=True)
    año_creacion = models.CharField(max_length=20, blank=True, null=True)
    activos = models.DecimalField(max_digits=15, decimal_places=2, blank=True, null=True)
    cant_bienes = models.IntegerField(blank=True, null=True)
    cant_bancos = models.IntegerField(blank=True, null=True)
    cant_cuentas = models.IntegerField(blank=True, null=True)
    cant_inversiones = models.IntegerField(blank=True, null=True)
    pasivos = models.DecimalField(max_digits=15, decimal_places=2, blank=True, null=True)
    cant_deudas = models.IntegerField(blank=True, null=True)
    patrimonio = models.DecimalField(max_digits=15, decimal_places=2, blank=True, null=True)
    apalancamiento = models.DecimalField(max_digits=10, decimal_places=2, blank=True, null=True)
    endeudamiento = models.DecimalField(max_digits=10, decimal_places=2, blank=True, null=True)
    aum_pat_subito = models.CharField(max_length=50, blank=True, null=True)
    activos_var_abs = models.CharField(max_length=50, blank=True, null=True)
    activos_var_rel = models.CharField(max_length=50, blank=True, null=True)
    pasivos_var_abs = models.CharField(max_length=50, blank=True, null=True)
    pasivos_var_rel = models.CharField(max_length=50, blank=True, null=True)
    patrimonio_var_abs = models.CharField(max_length=50, blank=True, null=True)
    patrimonio_var_rel = models.CharField(max_length=50, blank=True, null=True)
    apalancamiento_var_abs = models.CharField(max_length=50, blank=True, null=True)
    apalancamiento_var_rel = models.CharField(max_length=50, blank=True, null=True)
    endeudamiento_var_abs = models.CharField(max_length=50, blank=True, null=True)
    endeudamiento_var_rel = models.CharField(max_length=50, blank=True, null=True)
    banco_saldo = models.DecimalField(max_digits=15, decimal_places=2, blank=True, null=True)
    bienes = models.DecimalField(max_digits=15, decimal_places=2, blank=True, null=True)
    inversiones = models.DecimalField(max_digits=15, decimal_places=2, blank=True, null=True)
    banco_saldo_var_abs = models.CharField(max_length=50, blank=True, null=True)
    banco_saldo_var_rel = models.CharField(max_length=50, blank=True, null=True)
    bienes_var_abs = models.CharField(max_length=50, blank=True, null=True)
    bienes_var_rel = models.CharField(max_length=50, blank=True, null=True)
    inversiones_var_abs = models.CharField(max_length=50, blank=True, null=True)
    inversiones_var_rel = models.CharField(max_length=50, blank=True, null=True)
    ingresos = models.DecimalField(max_digits=15, decimal_places=2, blank=True, null=True)
    cant_ingresos = models.IntegerField(blank=True, null=True)
    ingresos_var_abs = models.CharField(max_length=50, blank=True, null=True)
    ingresos_var_rel = models.CharField(max_length=50, blank=True, null=True)
    capital = models.DecimalField(max_digits=15, decimal_places=2, blank=True, null=True)
    last_updated = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Reporte Financiero"
        verbose_name_plural = "Reportes Financieros"
        ordering = ['-ano_declaracion']

    def __str__(self):
        return f"Reporte de {self.person.nombre_completo} ({self.ano_declaracion})"
    
class Conflict(models.Model):
    person = models.ForeignKey(Person, on_delete=models.CASCADE, related_name='conflicts')
    fecha_inicio = models.DateField(verbose_name="Fecha de Inicio", null=True, blank=True)
    q1 = models.BooleanField(verbose_name="Accionista de algún proveedor del grupo", default=False)
    q2 = models.BooleanField(verbose_name="Familiar accionista, proveedor, empleado", default=False)
    q3 = models.BooleanField(verbose_name="Accionista de alguna compania del grupo", default=False)
    q4 = models.BooleanField(verbose_name="Actividades extralaborales", default=False)
    q5 = models.BooleanField(verbose_name="Negocios o bienes con empleados del grupo", default=False)
    q6 = models.BooleanField(verbose_name="Participación en juntas o consejos directivos", default=False)
    q7 = models.BooleanField(verbose_name="Potencial conflicto diferente a los anteriores", default=False)
    q8 = models.BooleanField(verbose_name="Consciente del código de conducta empresarial", default=False)
    q9 = models.BooleanField(verbose_name="Veracidad de la información consignada", default=False)
    q10 = models.BooleanField(verbose_name="Familiar de funcionario público", default=False)
    q11 = models.BooleanField(verbose_name="Relacion con el sector o funcionario público", default=False)
    last_updated = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Conflicto"
        verbose_name_plural = "Conflictos"

    def __str__(self):
        return f"Conflictos de {self.person.nombre_completo}"

class Card(models.Model):
    CARD_TYPE_CHOICES = [
        ('MC', 'Mastercard'),
        ('VI', 'Visa'),
    ]
    
    person = models.ForeignKey(Person, on_delete=models.CASCADE, related_name='cards')
    card_type = models.CharField(max_length=2, choices=CARD_TYPE_CHOICES)
    card_number = models.CharField(max_length=20)
    transaction_date = models.DateField()
    description = models.TextField()
    original_value = models.DecimalField(max_digits=15, decimal_places=2)
    exchange_rate = models.DecimalField(max_digits=10, decimal_places=4, null=True, blank=True)
    charges = models.DecimalField(max_digits=15, decimal_places=2, null=True, blank=True)
    balance = models.DecimalField(max_digits=15, decimal_places=2, null=True, blank=True)
    installments = models.CharField(max_length=20, null=True, blank=True)
    source_file = models.CharField(max_length=255)
    page_number = models.IntegerField()
    last_updated = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Tarjeta"
        verbose_name_plural = "Tarjetas"
        ordering = ['-transaction_date']

    def __str__(self):
        return f"{self.get_card_type_display()} - {self.card_number} - {self.transaction_date}"
"@

# Create views.py with import functionality
Set-Content -Path "core/views.py" -Value @"
from django.http import HttpResponse, HttpResponseRedirect
from django.template import loader
from django.shortcuts import render
from .models import Person, FinancialReport, Conflict, Card
import pandas as pd
from django.contrib import messages
from django.core.paginator import Paginator
from django.db.models import Q
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
import os
import glob
from django.contrib.auth.decorators import login_required

@login_required
def main(request):
    persons = Person.objects.all()
    persons = _apply_person_filters_and_sorting(persons, request.GET)
    
    if 'export' in request.GET:
        model_fields = ['cedula', 'nombre_completo', 'cargo', 'correo', 'compania', 'estado', 'revisar', 'comments']
        return export_to_excel(persons, model_fields, 'persons_export')

    """
    Main view showing the list of persons with filtering and pagination
    """
    persons = Person.objects.all()
    persons = _apply_person_filters_and_sorting(persons, request.GET)
    
    # Get dropdown values
    dropdown_values = _get_dropdown_values()
    
    # Pagination
    paginator = Paginator(persons, 1000)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)
    
    context = {
        'page_obj': page_obj,
        'persons': page_obj.object_list,
        'persons_count': persons.count(),
        'current_order': request.GET.get('order_by', 'nombre_completo').replace('-', ''),
        'current_direction': request.GET.get('sort_direction', 'asc'),
        'all_params': {k: v for k, v in request.GET.items() if k not in ['order_by', 'sort_direction']},
        **dropdown_values
    }
    return render(request, 'persons.html', context)
"@

    # Create admin.py with enhanced configuration
Set-Content -Path "core/admin.py" -Value @" 
from django.contrib import admin
from .models import Person, FinancialReport, Conflict, Card

def make_active(modeladmin, request, queryset):
    queryset.update(estado='Activo')
make_active.short_description = "Mark selected as Active"

def make_retired(modeladmin, request, queryset):
    queryset.update(estado='Retirado')
make_retired.short_description = "Mark selected as Retired"

def mark_for_check(modeladmin, request, queryset):
    queryset.update(revisar=True)
mark_for_check.short_description = "Mark for check"

def unmark_for_check(modeladmin, request, queryset):
    queryset.update(revisar=False)
unmark_for_check.short_description = "Unmark for check"

class PersonAdmin(admin.ModelAdmin):
    list_display = ("cedula", "nombre_completo", "cargo", "correo", "compania", "estado", "revisar")
    search_fields = ("nombre_completo", "cedula", "comments")
    list_filter = ("estado", "compania", "revisar")
    list_per_page = 25
    ordering = ('nombre_completo',)
    actions = [make_active, make_retired, mark_for_check, unmark_for_check]
    
    fieldsets = (
        (None, {
            'fields': ('cedula', 'nombre_completo', 'cargo')
        }),
        ('Advanced options', {
            'classes': ('collapse',),
            'fields': ('correo', 'compania', 'estado', 'revisar', 'comments'),
        }),
    )
    
class FinancialReportAdmin(admin.ModelAdmin):
    list_display = ('person', 'ano_declaracion', 'patrimonio', 'activos', 'pasivos', 'last_updated')
    list_filter = ('ano_declaracion', 'person__compania', 'person__estado')
    search_fields = ('person__nombre_completo', 'person__cedula')
    list_per_page = 25
    raw_id_fields = ('person',)
    
    fieldsets = (
        (None, {
            'fields': ('person', 'ano_declaracion', 'año_creacion')
        }),
        ('Financial Data', {
            'fields': (
                ('activos', 'pasivos', 'patrimonio'),
                ('apalancamiento', 'endeudamiento'),
                ('banco_saldo', 'bienes', 'inversiones'),
                ('ingresos', 'cant_ingresos'),
                ('aum_pat_subito', 'capital')
            )
        }),
        ('Variations', {
            'classes': ('collapse',),
            'fields': (
                ('activos_var_abs', 'activos_var_rel'),
                ('pasivos_var_abs', 'pasivos_var_rel'),
                ('patrimonio_var_abs', 'patrimonio_var_rel'),
                ('apalancamiento_var_abs', 'apalancamiento_var_rel'),
                ('endeudamiento_var_abs', 'endeudamiento_var_rel'),
                ('banco_saldo_var_abs', 'banco_saldo_var_rel'),
                ('bienes_var_abs', 'bienes_var_rel'),
                ('inversiones_var_abs', 'inversiones_var_rel'),
                ('ingresos_var_abs', 'ingresos_var_rel')
            )
        })
    )

class ConflictAdmin(admin.ModelAdmin):
    list_display = ('person', 'fecha_inicio', 'q1', 'q2', 'q3', 'q4', 'q5', 'q6', 'q7', 'q8', 'q9', 'q10', 'q11')
    list_filter = ('q1', 'q2', 'q3', 'q4', 'q5', 'q6', 'q7', 'q8', 'q9', 'q10', 'q11')
    search_fields = ('person__nombre_completo', 'person__cedula')
    raw_id_fields = ('person',)
    list_per_page = 25

class CardAdmin(admin.ModelAdmin):
    list_display = ('person', 'get_card_type_display', 'transaction_date', 'description', 'original_value')
    list_filter = ('card_type', 'transaction_date')
    search_fields = ('person__nombre_completo', 'person__cedula', 'description')
    date_hierarchy = 'transaction_date'
    list_per_page = 50

admin.site.register(Person, PersonAdmin)
admin.site.register(FinancialReport, FinancialReportAdmin)
admin.site.register(Conflict, ConflictAdmin)
admin.site.register(Card, CardAdmin)
"@

# Create urls.py for core app
Set-Content -Path "core/urls.py" -Value @"
from django.contrib.auth import views as auth_views
from django.urls import path
from . import views
from django.contrib.auth import get_user_model
from django.contrib import messages
from django.shortcuts import render, redirect

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
    <link rel="stylesheet" href="{% static 'css/freeze.css' %}">
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
            <a href="/persons/import/" class="btn btn-custom-primary" title="Importar">
                <i class="fas fa-database"></i>
            </a>
            <form method="post" action="{% url 'logout' %}" class="d-inline">
                {% csrf_token %}
                <button type="submit" class="btn btn-custom-primary" title="Cerrar sesión">
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

# Create login template
@"
{% extends "master.html" %}

{% block title %}Acceder{% endblock %}
{% block navbar_title %}Acceder{% endblock %}

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
                                <a href="{% url 'login' %}" class="btn btn-custom-primary" title="Recupera tu acceso">
                                    <i class="fas fa-sign-in-alt" style="color: rgb(0, 0, 255);"></i>
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