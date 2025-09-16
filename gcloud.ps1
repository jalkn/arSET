function arpa {
    param (
        [string]$ExcelFilePath = $null
    )

    $YELLOW = [ConsoleColor]::Yellow
    $GREEN = [ConsoleColor]::Green

    Write-Host "🚀 Creating ARPA" -ForegroundColor $YELLOW

    # Create python3 virtual environment
    python3 -m venv .venv
    .\.venv\scripts\activate

    # Install required python3 packages
    python3 -m pip install --upgrade pip
    python3 -m pip install django whitenoise django-bootstrap-v5 xlsxwriter openpyxl pandas xlrd>=2.0.1 pdfplumber PyMuPDF msoffcrypto-tool fuzzywuzzy python-Levenshtein

    # Create Django project
    django-admin startproject arpa
    cd arpa

    # Create core app
    python3 manage.py startapp core

    # Create templates directory structure
    $directories = @(
        "core/src",
        "core/static",
        "core/static/css",
        "core/static/js",
        "core/templates",
        "core/templatetags",
        "core/templates/admin",
        "core/templates/registration"
    )
    foreach ($dir in $directories) {
        New-Item -Path $dir -ItemType Directory -Force
    }

# Create models.py with cedula as primary key
Set-Content -Path "core/models.py" -Value @" 
from django.db import models
from django.contrib.auth.models import User

class Person(models.Model):
    cedula = models.CharField(max_length=20, primary_key=True)
    nombre_completo = models.CharField(max_length=255)
    correo = models.EmailField(max_length=255, blank=True)
    estado = models.CharField(max_length=50, default='Activo')
    compania = models.CharField(max_length=255, blank=True)
    cargo = models.CharField(max_length=255, blank=True)
    area = models.CharField(max_length=255, blank=True)
    revisar = models.BooleanField(default=False)
    comments = models.TextField(blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f"{self.nombre_completo} ({self.cedula})"

class Conflict(models.Model):
    id = models.AutoField(primary_key=True)
    person = models.ForeignKey(
        Person,
        on_delete=models.CASCADE,
        related_name='conflicts',
        to_field='cedula',
        db_column='cedula'
    )
    ano = models.IntegerField(null=True, blank=True)
    fecha_inicio = models.DateField(null=True, blank=True)
    q1 = models.BooleanField(null=True, blank=True) 
    q1_detalle = models.TextField(blank=True)
    q2 = models.BooleanField(null=True, blank=True) 
    q2_detalle = models.TextField(blank=True)
    q3 = models.BooleanField(null=True, blank=True) 
    q3_detalle = models.TextField(blank=True)
    q4 = models.BooleanField(null=True, blank=True) 
    q4_detalle = models.TextField(blank=True)
    q5 = models.BooleanField(null=True, blank=True) 
    q5_detalle = models.TextField(blank=True)
    q6 = models.BooleanField(null=True, blank=True) 
    q6_detalle = models.TextField(blank=True)
    q7 = models.BooleanField(null=True, blank=True) 
    q7_detalle = models.TextField(blank=True)
    q8 = models.BooleanField(null=True, blank=True) 
    q9 = models.BooleanField(null=True, blank=True) 
    q10 = models.BooleanField(null=True, blank=True) 
    q10_detalle = models.TextField(blank=True)
    q11 = models.BooleanField(null=True, blank=True) 
    q11_detalle = models.TextField(blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f"Conflictos para {self.person.nombre_completo} (ID: {self.id}, Año: {self.ano})"

class FinancialReport(models.Model):
    id = models.AutoField(primary_key=True)
    person = models.ForeignKey(
        Person,
        on_delete=models.CASCADE,
        related_name='financial_reports',
        to_field='cedula',
        db_column='cedula'
    )
    fk_id_periodo = models.IntegerField(null=True, blank=True)
    ano_declaracion = models.IntegerField(null=True, blank=True)
    ano_creacion = models.IntegerField(null=True, blank=True)
    activos = models.FloatField(null=True, blank=True)
    cant_bienes = models.IntegerField(null=True, blank=True)
    cant_bancos = models.IntegerField(null=True, blank=True)
    cant_cuentas = models.IntegerField(null=True, blank=True)
    cant_inversiones = models.IntegerField(null=True, blank=True)
    pasivos = models.FloatField(null=True, blank=True)
    cant_deudas = models.IntegerField(null=True, blank=True)
    patrimonio = models.FloatField(null=True, blank=True)
    apalancamiento = models.FloatField(null=True, blank=True)
    endeudamiento = models.FloatField(null=True, blank=True)
    capital = models.FloatField(null=True, blank=True)
    aum_pat_subito = models.FloatField(null=True, blank=True)
    activos_var_abs = models.FloatField(null=True, blank=True)
    activos_var_rel = models.CharField(max_length=50, null=True, blank=True)
    pasivos_var_abs = models.FloatField(null=True, blank=True)
    pasivos_var_rel = models.CharField(max_length=50, null=True, blank=True)
    patrimonio_var_abs = models.FloatField(null=True, blank=True)
    patrimonio_var_rel = models.CharField(max_length=50, null=True, blank=True)
    apalancamiento_var_abs = models.FloatField(null=True, blank=True)
    apalancamiento_var_rel = models.CharField(max_length=50, null=True, blank=True)
    endeudamiento_var_abs = models.FloatField(null=True, blank=True)
    endeudamiento_var_rel = models.CharField(max_length=50, null=True, blank=True)
    banco_saldo = models.FloatField(null=True, blank=True)
    bienes = models.FloatField(null=True, blank=True)
    inversiones = models.FloatField(null=True, blank=True)
    banco_saldo_var_abs = models.FloatField(null=True, blank=True)
    banco_saldo_var_rel = models.CharField(max_length=50, null=True, blank=True)
    bienes_var_abs = models.FloatField(null=True, blank=True)
    bienes_var_rel = models.CharField(max_length=50, null=True, blank=True)
    inversiones_var_abs = models.FloatField(null=True, blank=True)
    inversiones_var_rel = models.CharField(max_length=50, null=True, blank=True)
    ingresos = models.FloatField(null=True, blank=True)
    cant_ingresos = models.IntegerField(null=True, blank=True)
    ingresos_var_abs = models.FloatField(null=True, blank=True)
    ingresos_var_rel = models.CharField(max_length=50, null=True, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f"Reporte Financiero para {self.person.nombre_completo} (Periodo: {self.fk_id_periodo})"
    

class CreditCard(models.Model):
    id = models.AutoField(primary_key=True)

    person = models.ForeignKey(
        Person,
        on_delete=models.SET_NULL,
        related_name='credit_cards', 
        to_field='cedula',
        db_column='cedula', 
        null=True,
        blank=True
    )

    tipo_tarjeta = models.CharField(max_length=50, null=True, blank=True)
    numero_tarjeta = models.CharField(max_length=20, null=True, blank=True) # Número de Tarjeta
    moneda = models.CharField(max_length=10, null=True, blank=True)
    trm_cierre = models.CharField(max_length=50, null=True, blank=True) # Renombrado y tipo ajustado a cómo lo genera tcs.py (string)
    valor_original = models.CharField(max_length=50, null=True, blank=True) # Tipo ajustado a string
    valor_cop = models.CharField(max_length=50, null=True, blank=True) # Agregado de nuevo, tipo ajustado a string
    numero_autorizacion = models.CharField(max_length=100, null=True, blank=True)
    fecha_transaccion = models.DateField(null=True, blank=True)
    dia = models.CharField(max_length=20, null=True, blank=True)
    descripcion = models.TextField(null=True, blank=True)
    categoria = models.CharField(max_length=255, null=True, blank=True)
    subcategoria = models.CharField(max_length=255, null=True, blank=True)
    zona = models.CharField(max_length=255, null=True, blank=True)

    def __str__(self):
        return f"{self.descripcion} - {self.valor_cop} (Tarjeta: {self.numero_tarjeta})"

    class Meta:
        verbose_name = "Tarjeta de Crédito"
        verbose_name_plural = "Tarjetas de Crédito"
        unique_together = ('person', 'fecha_transaccion', 'numero_autorizacion', 'valor_original')
"@

# Create admin.py with enhanced configuration
Set-Content -Path "core/admin.py" -Value @" 
from django.contrib import admin
from django import forms
from django.utils.html import format_html
from django.urls import reverse
from core.models import Person, Conflict, FinancialReport

class ConflictForm(forms.ModelForm):
    class Meta:
        model = Conflict
        fields = '__all__'

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Replace boolean field widgets with custom display
        for field_name in ['q1', 'q2', 'q3', 'q4', 'q5', 'q6', 'q7', 'q8', 'q9', 'q10', 'q11']:
            self.fields[field_name].widget = forms.Select(choices=[(True, 'YES'), (False, 'NO')])
        # Make detail fields optional as they might not always be filled
        for field_name in ['q1_detalle', 'q2_detalle', 'q3_detalle', 'q4_detalle', 'q5_detalle',
                           'q6_detalle', 'q7_detalle', 'q10_detalle', 'q11_detalle']: # New detail fields
            if field_name in self.fields: # Check if field exists to prevent errors if not added in models
                self.fields[field_name].required = False # New field


@admin.register(Person)
class PersonAdmin(admin.ModelAdmin):
    list_display = ('cedula', 'nombre_completo', 'cargo', 'area', 'compania', 'estado', 'revisar')
    search_fields = ('cedula', 'nombre_completo', 'correo')
    list_filter = ('estado', 'compania', 'revisar')
    list_editable = ('revisar',)

    # Custom fields to show in detail view
    readonly_fields = ('cedula_with_actions', 'conflicts_link', 'financial_reports_link')

    fieldsets = (
        (None, {
            'fields': ('cedula_with_actions', 'nombre_completo', 'correo', 'estado', 'compania', 'cargo', 'area', 'revisar', 'comments')
        }),
        ('Related Records', {
            'fields': ('conflicts_link', 'financial_reports_link'),
            'classes': ('collapse',)
        }),
    )

    def cedula_with_actions(self, obj):
        if obj.pk:
            change_url = reverse('admin:core_person_change', args=[obj.pk])
            history_url = reverse('admin:core_person_history', args=[obj.pk])
            add_url = reverse('admin:core_person_add')

            return format_html(
                '{} <div class="nowrap">'
                '<a href="{}" class="changelink">Change</a> &nbsp;'
                '<a href="{}" class="historylink">History</a> &nbsp;'
                '<a href="{}" class="addlink">Add another</a>'
                '</div>',
                obj.cedula,
                change_url,
                history_url,
                add_url
            )
        return obj.cedula
    cedula_with_actions.short_description = 'Cedula'

    def conflicts_link(self, obj):
        if obj.pk:
            conflict = obj.conflicts.first()
            if conflict:
                change_url = reverse('admin:core_conflict_change', args=[conflict.pk])
                add_url = reverse('admin:core_conflict_add') + f'?person={obj.pk}'
                list_url = reverse('admin:core_conflict_changelist') + f'?q={obj.cedula}'

                return format_html(
                    '<div class="nowrap">'
                    '<a href="{}" class="changelink">View/Edit Conflicts</a> &nbsp;'
                    '<a href="{}" class="addlink">Add New Conflict</a> &nbsp;'
                    '<a href="{}" class="viewlink">All Conflicts</a>'
                    '</div>',
                    change_url,
                    add_url,
                    list_url
                )
            else:
                add_url = reverse('admin:core_conflict_add') + f'?person={obj.pk}'
                return format_html(
                    '<a href="{}" class="addlink">Create Conflict Record</a>',
                    add_url
                )
        return "-"
    conflicts_link.short_description = 'Conflict Records'
    conflicts_link.allow_tags = True

    def financial_reports_link(self, obj):
        if obj.pk:
            report = obj.financial_reports.first()
            if report:
                change_url = reverse('admin:core_financialreport_change', args=[report.pk])
                add_url = reverse('admin:core_financialreport_add') + f'?person={obj.pk}'
                list_url = reverse('admin:core_financialreport_changelist') + f'?q={obj.cedula}'

                return format_html(
                    '<div class="nowrap">'
                    '<a href="{}" class="changelink">View/Editar Declaracion B&R</a> &nbsp;'
                    '<a href="{}" class="addlink">Agregar Nueva declaracion B&R</a> &nbsp;'
                    '<a href="{}" class="viewlink">Todo en Bienes y Rentas</a>'
                    '</div>',
                    change_url,
                    add_url,
                    list_url
                )
            else:
                add_url = reverse('admin:core_financialreport_add') + f'?person={obj.pk}'
                return format_html(
                    '<a href="{}" class="addlink">Create Financial Report Record</a>',
                    add_url
                )
        return "-"
    financial_reports_link.short_description = 'Bienes y Rentas'
    financial_reports_link.allow_tags = True

    def get_fieldsets(self, request, obj=None):
        if obj is None:  # Add view
            return [(None, {'fields': ('cedula', 'nombre_completo', 'correo', 'estado', 'compania', 'cargo', 'revisar', 'comments')})]
        return super().get_fieldsets(request, obj)

@admin.register(Conflict)
class ConflictAdmin(admin.ModelAdmin):
    form = ConflictForm
    # Add new detail fields to list_display
    list_display = ('person', 'fecha_inicio',
                    'get_q1_display', 'q1_detalle',
                    'get_q2_display', 'q2_detalle',
                    'get_q3_display', 'q3_detalle',
                    'get_q4_display', 'q4_detalle',
                    'get_q5_display', 'q5_detalle',
                    'get_q6_display', 'q6_detalle',
                    'get_q7_display', 'q7_detalle',
                    'get_q8_display', 'get_q9_display',
                    'get_q10_display', 'q10_detalle',
                    'get_q11_display', 'q11_detalle')
    # Add new detail fields to list_filter and search_fields if desired
    list_filter = ('q1', 'q2', 'q3', 'q4', 'q5', 'q6', 'q7', 'q8', 'q9', 'q10', 'q11')
    search_fields = ('person__nombre_completo', 'person__cedula',
                     'q1_detalle', 'q2_detalle', 'q3_detalle', 'q4_detalle', # New search fields
                     'q5_detalle', 'q6_detalle', 'q7_detalle', 'q10_detalle', 'q11_detalle') # New search fields
    raw_id_fields = ('person',)

    # Update fieldsets to include new detail fields
    fieldsets = (
        (None, {
            'fields': ('person', 'fecha_inicio')
        }),
        ('Conflict Questions', {
            'fields': (
                ('q1', 'q1_detalle'), # Group boolean with its detail
                ('q2', 'q2_detalle'), # Group boolean with its detail
                ('q3', 'q3_detalle'), # Group boolean with its detail
                ('q4', 'q4_detalle'), # Group boolean with its detail
                ('q5', 'q5_detalle'), # Group boolean with its detail
                ('q6', 'q6_detalle'), # Group boolean with its detail
                ('q7', 'q7_detalle'), # Group boolean with its detail
                'q8', 'q9',
                ('q10', 'q10_detalle'), # Group boolean with its detail
                ('q11', 'q11_detalle'), # Group boolean with its detail
            ),
            'description': 'Answer "YES" or "NO" to each question and provide details where applicable'
        }),
    )

    def get_form(self, request, obj=None, **kwargs):
        form = super().get_form(request, obj, **kwargs)
        form.base_fields['q1'].label = 'Accionista de proveedor'
        form.base_fields['q1_detalle'].label = 'Accionista de proveedor (Detalle)' # New label
        form.base_fields['q2'].label = 'Familiar de accionista/empleado'
        form.base_fields['q2_detalle'].label = 'Familiar de accionista/empleado (Detalle)' # New label
        form.base_fields['q3'].label = 'Accionista del grupo'
        form.base_fields['q3_detalle'].label = 'Accionista del grupo (Detalle)' # New label
        form.base_fields['q4'].label = 'Actividades extralaborales'
        form.base_fields['q4_detalle'].label = 'Actividades extralaborales (Detalle)' # New label
        form.base_fields['q5'].label = 'Negocios con empleados'
        form.base_fields['q5_detalle'].label = 'Negocios con empleados (Detalle)' # New label
        form.base_fields['q6'].label = 'Participacion en juntas'
        form.base_fields['q6_detalle'].label = 'Participacion en juntas (Detalle)' # New label
        form.base_fields['q7'].label = 'Otro conflicto'
        form.base_fields['q7_detalle'].label = 'Otro conflicto (Detalle)' # New label
        form.base_fields['q8'].label = 'Conoce codigo de conducta'
        form.base_fields['q9'].label = 'Veracidad de informacion'
        form.base_fields['q10'].label = 'Familiar de funcionario'
        form.base_fields['q10_detalle'].label = 'Familiar de funcionario (Detalle)' # New label
        form.base_fields['q11'].label = 'Relacion con sector publico'
        form.base_fields['q11_detalle'].label = 'Relacion con sector publico (Detalle)' # New label
        return form

    # YES/NO display methods for list view (no changes here for detail fields)
    def get_q1_display(self, obj): return "YES" if obj.q1 else "NO"
    get_q1_display.short_description = 'Accionista de proveedor'
    def get_q2_display(self, obj): return "YES" if obj.q2 else "NO"
    get_q2_display.short_description = 'Familiar de accionista/empleado'
    def get_q3_display(self, obj): return "YES" if obj.q3 else "NO"
    get_q3_display.short_description = 'Accionista del grupo'
    def get_q4_display(self, obj): return "YES" if obj.q4 else "NO"
    get_q4_display.short_description = 'Actividades extralaborales'
    def get_q5_display(self, obj): return "YES" if obj.q5 else "NO"
    get_q5_display.short_description = 'Negocios con empleados'
    def get_q6_display(self, obj): return "YES" if obj.q6 else "NO"
    get_q6_display.short_description = 'Participacion en juntas'
    def get_q7_display(self, obj): return "YES" if obj.q7 else "NO"
    get_q7_display.short_description = 'Otro conflicto'
    def get_q8_display(self, obj): return "YES" if obj.q8 else "NO"
    get_q8_display.short_description = 'Conoce codigo de conducta'
    def get_q9_display(self, obj): return "YES" if obj.q9 else "NO"
    get_q9_display.short_description = 'Veracidad de informacion'
    def get_q10_display(self, obj): return "YES" if obj.q10 else "NO"
    get_q10_display.short_description = 'Familiar de funcionario'
    def get_q11_display(self, obj): return "YES" if obj.q11 else "NO"
    get_q11_display.short_description = 'Relacion con sector publico'

@admin.register(FinancialReport) # Register the new model
class FinancialReportAdmin(admin.ModelAdmin):
    list_display = (
        'person', 'fk_id_periodo', 'ano_declaracion', 'activos', 'pasivos',
        'patrimonio', 'ingresos', 'apalancamiento', 'endeudamiento',
        'activos_var_rel', 'pasivos_var_rel', 'patrimonio_var_rel',
        'ingresos_var_rel'
    )
    search_fields = (
        'person__nombre_completo', 'person__cedula', 'fk_id_periodo',
        'ano_declaracion'
    )
    list_filter = ('ano_declaracion', 'fk_id_periodo')
    raw_id_fields = ('person',)

    fieldsets = (
        (None, {
            'fields': ('person', 'fk_id_periodo', 'ano_declaracion', 'ano_creacion')
        }),
        ('Financial Data', {
            'fields': (
                'activos', 'cant_bienes', 'cant_bancos', 'cant_cuentas', 'cant_inversiones',
                'pasivos', 'cant_deudas', 'patrimonio', 'capital', 'aum_pat_subito',
                'banco_saldo', 'bienes', 'inversiones', 'ingresos', 'cant_ingresos'
            )
        }),
        ('Trends and Variations', {
            'fields': (
                ('apalancamiento', 'apalancamiento_var_abs', 'apalancamiento_var_rel'),
                ('endeudamiento', 'endeudamiento_var_abs', 'endeudamiento_var_rel'),
                ('activos_var_abs', 'activos_var_rel'),
                ('pasivos_var_abs', 'pasivos_var_rel'),
                ('patrimonio_var_abs', 'patrimonio_var_rel'),
                ('banco_saldo_var_abs', 'banco_saldo_var_rel'),
                ('bienes_var_abs', 'bienes_var_rel'),
                ('inversiones_var_abs', 'inversiones_var_rel'),
                ('ingresos_var_abs', 'ingresos_var_rel'),
            )
        }),
    )
"@

# Create urls.py for core app
Set-Content -Path "core/urls.py" -Value @"
# urls.py
from django.contrib.auth import views as auth_views
from django.urls import path
from . import views
from django.contrib.auth import get_user_model
from django.contrib import messages
from django.shortcuts import render, redirect
from django.urls import path
from django.contrib.auth import views as auth_views
from .views import (main, register_superuser, ImportView, person_list,
                   import_conflicts, conflict_list, import_persons,
                   import_finances, person_details, financial_report_list,
                   export_persons_excel, alerts_list, save_comment, delete_comment, import_tcs, import_categorias,
                   tcs_list) 

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
    path('', main, name='main'),
    path('login/', auth_views.LoginView.as_view(template_name='registration/login.html'), name='login'),
    path('logout/', auth_views.LogoutView.as_view(next_page='login'), name='logout'),
    path('register/', register_superuser, name='register'),
    path('import/', ImportView.as_view(), name='import'),
    path('import/persons/', import_persons, name='import_persons'),
    path('import/conflicts/', import_conflicts, name='import_conflicts'),
    path('import/finances/', import_finances, name='import_finances'),
    path('import/tcs/', import_tcs, name='import_tcs'),
    path('import/categorias/', import_categorias, name='import_categorias'),
    path('persons/', person_list, name='person_list'),
    path('persons/<str:cedula>/', person_details, name='person_details'),
    path('persons/export/excel/', export_persons_excel, name='export_persons_excel'),
    path('financial_reports/', financial_report_list, name='financial_report_list'),
    path('alerts/', alerts_list, name='alerts_list'),
    path('persons/<str:cedula>/toggle_revisar/', views.toggle_revisar_status, name='toggle_revisar_status'),
    path('persons/<str:cedula>/save_comment/', save_comment, name='save_comment'),
    path('persons/<str:cedula>/delete_comment/', delete_comment, name='delete_comment'),
    path('conflicts/', conflict_list, name='conflict_list'),
    path('tcs/', tcs_list, name='tcs_list'), 
]
"@

# Update core/views.py with financial import
Set-Content -Path "core/views.py" -Value @"
import pandas as pd
from datetime import datetime
import os
from django.conf import settings
from django.http import HttpResponseRedirect, HttpResponse
from django.shortcuts import render, redirect
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.contrib.auth import get_user_model
from django.contrib.auth.mixins import LoginRequiredMixin
from django.views.generic import TemplateView
from django.core.paginator import Paginator
from core.models import Person, Conflict, FinancialReport, CreditCard
from django.db.models import Q
import subprocess
import msoffcrypto
import io
import re
from django.views.decorators.http import require_POST
from django.shortcuts import get_object_or_404, redirect
from . import tcs 

@login_required
@require_POST
def toggle_revisar_status(request, cedula):
    """
    Toggles the 'revisar' status for a given person.
    Expects a POST request with the person's cedula.
    """
    person = get_object_or_404(Person, cedula=cedula)
    person.revisar = not person.revisar  # Toggle the boolean value
    person.save()

    messages.success(request, f"Revisar status for {person.nombre_completo} ({person.cedula}) updated successfully.")

    # Redirect back to the page the request came from
    next_url = request.META.get('HTTP_REFERER')
    if next_url:
        return redirect(next_url)
    else:
        return redirect('financial_report_list') # Or 'main' or 'alerts_list' as a default

# Helper function to clean and convert numeric values from strings
def _clean_numeric_value(value):
    if pd.isna(value):
        return None

    str_value = str(value).strip()
    if not str_value:
        return None

    numeric_part = re.sub(r'[^\d.%\-]', '', str_value)

    try:
        if '%' in numeric_part:
            # If it's a percentage, convert to float and divide by 100
            return float(numeric_part.replace('%', '')) / 100
        else:
            # Otherwise, just convert to float
            return float(numeric_part)
    except ValueError:
        return None # Return None if conversion fails

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
        """
        Overrides the default get_context_data to add counts for persons and conflicts,
        and to gather analysis results from the core/src directory.
        """
        context = super().get_context_data(**kwargs)
        # These counts are fetched from models directly
        context['conflict_count'] = Conflict.objects.count()
        context['person_count'] = Person.objects.count()
        context['finances_count'] = FinancialReport.objects.count()
        context['alerts_count'] = Person.objects.filter(revisar=True).count()
        # Updated to count CreditCard entries
        context['tc_count'] = CreditCard.objects.count()


        analysis_results = []
        core_src_dir = os.path.join(settings.BASE_DIR, 'core', 'src')

        # Helper function to get file status
        def get_file_status(filename, directory=core_src_dir):
            file_path = os.path.join(directory, filename)
            status_info = {'filename': filename, 'records': '-', 'status': 'pending', 'last_updated': None, 'error': None}
            if os.path.exists(file_path):
                try:
                    df = pd.read_excel(file_path)
                    status_info['records'] = len(df)
                    status_info['status'] = 'success'
                    status_info['last_updated'] = datetime.fromtimestamp(os.path.getmtime(file_path))
                except Exception as e:
                    status_info['status'] = 'error'
                    status_info['error'] = f"Error reading file: {str(e)}"
            return status_info

        # --- Status for Personas.xlsx ---
        personas_status = get_file_status('Personas.xlsx')
        analysis_results.append(personas_status)
        if personas_status['status'] == 'success':
            context['person_count'] = personas_status['records']

        # --- Status for conflicts.xlsx ---
        conflicts_status = get_file_status('conflicts.xlsx')
        analysis_results.append(conflicts_status)
        if conflicts_status['status'] == 'success':
            context['conflict_count'] = conflicts_status['records']

        # --- Status for tcs.xlsx ---
        tcs_excel_status = get_file_status('tcs.xlsx')
        analysis_results.append(tcs_excel_status)
        if tcs_excel_status['status'] == 'success':
            context['tc_count'] = tcs_excel_status['records'] # Update tcs_count in context

        # --- Status for categorias.xlsx ---
        categorias_status = get_file_status('categorias.xlsx')
        analysis_results.append(categorias_status)
        if categorias_status['status'] == 'success':
            context['categorias_count'] = categorias_status['records']
        else:
            context['categorias_count'] = 0


        # --- Status for Nets.py output files ---
        analysis_results.append(get_file_status('bankNets.xlsx'))
        analysis_results.append(get_file_status('debtNets.xlsx'))
        analysis_results.append(get_file_status('goodNets.xlsx'))
        analysis_results.append(get_file_status('incomeNets.xlsx'))
        analysis_results.append(get_file_status('investNets.xlsx'))
        analysis_results.append(get_file_status('assetNets.xlsx'))
        analysis_results.append(get_file_status('worthNets.xlsx'))

        # --- Status for Trends.py output files ---
        analysis_results.append(get_file_status('trends.xlsx'))

        # --- Status for idTrends.xlsx ---
        idtrends_status = get_file_status('idTrends.xlsx')
        analysis_results.append(idtrends_status)
        if idtrends_status['status'] == 'success':
            context['financial_report_count'] = idtrends_status['records']


        context['analysis_results'] = analysis_results
        return context

@login_required
def tcs_list(request):
    """
    Function-based view to display credit card transactions from tcs.xlsx.
    """
    context = {}
    core_src_dir = os.path.join(settings.BASE_DIR, 'core', 'src')
    tcs_excel_path = os.path.join(core_src_dir, 'tcs.xlsx')
    personas_excel_path = os.path.join(core_src_dir, 'Personas.xlsx') # Assuming Personas.xlsx is also in core/src

    transactions_list = []

    print(f"DEBUG: Checking for tcs.xlsx at {tcs_excel_path}")
    if os.path.exists(tcs_excel_path):
        try:
            # Read tcs.xlsx, forcing 'Cedula' column to be read as a string
            tcs_df = pd.read_excel(tcs_excel_path, dtype={'Cedula': str})
            print(f"DEBUG: tcs_df loaded. Shape: {tcs_df.shape}")
            print(f"DEBUG: tcs_df columns RAW: {tcs_df.columns.tolist()}")

            # Standardize tcs_df column names for internal use (remove accents, lowercase, replace spaces with underscores)
            tcs_df.columns = [col.strip().lower().replace(' ', '_').replace('á', 'a').replace('é', 'e').replace('í', 'i').replace('ó', 'o').replace('ú', 'u').replace('.', '') for col in tcs_df.columns]
            print(f"DEBUG: tcs_df columns STANDARDIZED: {tcs_df.columns.tolist()}")

            # Read Personas.xlsx (if exists)
            personas_df = pd.DataFrame()
            print(f"DEBUG: Checking for Personas.xlsx at {personas_excel_path}")
            if os.path.exists(personas_excel_path):
                try:
                    personas_df = pd.read_excel(personas_excel_path)
                    print(f"DEBUG: personas_df loaded. Shape: {personas_df.shape}")
                    print(f"DEBUG: personas_df columns RAW: {personas_df.columns.tolist()}")
                    personas_df.columns = [col.strip().lower().replace(' ', '_').replace('á', 'a').replace('é', 'e').replace('í', 'i').replace('ó', 'o').replace('ú', 'u').replace('.', '') for col in personas_df.columns]
                    print(f"DEBUG: personas_df columns STANDARDIZED: {personas_df.columns.tolist()}")
                    
                    if 'cedula' in personas_df.columns:
                        # Use the clean_cedula_format function for robust cleaning
                        personas_df['cedula'] = personas_df['cedula'].apply(tcs.clean_cedula_format)
                        print("DEBUG: Personas 'cedula' column cleaned and standardized.")
                    else:
                        print("WARNING: 'cedula' column not found in Personas.xlsx.")
                        
                except Exception as e:
                    messages.warning(request, f"Error loading Personas.xlsx for joining: {e}")
                    print(f"ERROR: Loading Personas.xlsx: {e}")
            else:
                messages.warning(request, "Personas.xlsx not found. Person details might be missing.")
                print("WARNING: Personas.xlsx not found.")

            # Apply the cleaning function to the 'cedula' column of tcs_df after reading
            if 'cedula' in tcs_df.columns:
                tcs_df['cedula'] = tcs_df['cedula'].apply(tcs.clean_cedula_format)
                print("DEBUG: tcs_df 'cedula' column cleaned and standardized.")
            else:
                print("WARNING: 'cedula' column not found in tcs_df. Cannot merge with Personas data.")
                # If 'cedula' is missing in tcs_df, proceed without person data merge
                merged_df = tcs_df.copy()
                merged_df['nombre_completo'] = None
                merged_df['cargo'] = None
                merged_df['compania'] = None
                merged_df['area'] = None
                print("DEBUG: Proceeding without merging Personas data due to missing 'cedula' in tcs_df.")

            # Perform merge if both DFs and 'cedula' column are present
            if not personas_df.empty and 'cedula' in tcs_df.columns and 'cedula' in personas_df.columns:
                # Select only the necessary columns from personas_df to avoid conflicts
                personas_cols_to_merge = ['cedula', 'nombre_completo', 'cargo', 'compania']
                # Filter to only include columns that actually exist in personas_df
                existing_personas_cols = [col for col in personas_cols_to_merge if col in personas_df.columns]

                merged_df = pd.merge(tcs_df, personas_df[existing_personas_cols],
                                     on='cedula', how='left', suffixes=('', '_person'))
                print(f"DEBUG: DataFrames merged successfully. Merged_df shape: {merged_df.shape}")
                print(f"DEBUG: Merged_df columns AFTER MERGE: {merged_df.columns.tolist()}")
                print(f"DEBUG: First 5 rows of merged_df:\n{merged_df.head()}")
            else:
                # If merge cannot happen, use tcs_df as is and add placeholder columns
                merged_df = tcs_df.copy()
                if 'nombre_completo' not in merged_df.columns: merged_df['nombre_completo'] = None
                if 'cargo' not in merged_df.columns: merged_df['cargo'] = None
                if 'compania' not in merged_df.columns: merged_df['compania'] = None
                if 'area' not in merged_df.columns: merged_df['area'] = None
                print("WARNING: Merge skipped (personas_df empty or 'cedula' column missing). Merged_df is a copy of tcs_df.")
                print(f"DEBUG: Merged_df (unmerged) columns: {merged_df.columns.tolist()}")


            # Map standardized Excel columns to dictionary keys for the template
            for index, row in merged_df.iterrows():
                # Robust date parsing
                fecha_transaccion_raw = row.get('fecha_de_transaccion', None)
                if pd.isna(fecha_transaccion_raw): # Check for pandas NaN values
                    fecha_transaccion = None
                else:
                    try:
                        # Try converting to datetime if it's already a datetime object or string
                        if isinstance(fecha_transaccion_raw, datetime):
                            fecha_transaccion = fecha_transaccion_raw.date() # Get only date part
                        else:
                            # Attempt to parse common date formats if it's a string
                            fecha_transaccion = pd.to_datetime(str(fecha_transaccion_raw)).date()
                    except ValueError:
                        fecha_transaccion = None # If parsing fails, set to None

                person_data = {
                    'cedula': row.get('cedula', ''), # Ensure cedula is always present for URL reversal
                    'nombre_completo': row.get('nombre_completo', 'N/A'),
                    'cargo': row.get('cargo', 'N/A'),
                    'compania': row.get('compania', 'N/A'),
                    'area': row.get('area', 'N/A'),
                }
                
                transaction = {
                    'person': person_data,
                    # Map standardized Excel column names (from tcs.xlsx columns list) to desired keys
                    'tipo_tarjeta': row.get('tipo_de_tarjeta', 'N/A'),
                    'numero_tarjeta': row.get('numero_de_tarjeta', 'N/A'),
                    'moneda': row.get('moneda', 'N/A'),
                    'trm_cierre': row.get('trm_cierre', 'N/A'),
                    'valor_original': row.get('valor_original', 'N/A'),
                    'valor_cop': row.get('valor_cop', 'N/A'),
                    'numero_autorizacion': row.get('numero_de_autorizacion', 'N/A'),
                    'fecha_transaccion': fecha_transaccion, # Use the parsed date
                    'dia': row.get('dia', 'N/A'),
                    'descripcion': row.get('descripcion', 'N/A'),
                    'categoria': row.get('categoria', 'N/A'),
                    'subcategoria': row.get('subcategoria', 'N/A'),
                    'zona': row.get('zona', 'N/A'),
                    'tasa_pactada': row.get('tasa_pactada', 'N/A'), # Added
                    'tasa_ea_facturada': row.get('tasa_ea_facturada', 'N/A'), # Added
                    'cargos_y_abonos': row.get('cargos_y_abonos', 'N/A'), # Added
                    'saldo_a_diferir': row.get('saldo_a_diferir', 'N/A'), # Added
                    'cuotas': row.get('cuotas', 'N/A'), # Added
                    'pagina': row.get('pagina', 'N/A'), # Added
                    'tar_x_per': row.get('tar_x_per', 'N/A'), # Added for 'Tar. x Per.'
                    'archivo': row.get('archivo', 'N/A'), # Added for 'Archivo'
                }
                transactions_list.append(transaction)

            # --- START FILTERING LOGIC ---
            q = request.GET.get('q')
            if q:
                query_lower = q.lower()
                filtered_list = []
                for transaction in transactions_list:
                    # Check if the query is in any of the relevant string values
                    if any(query_lower in str(value).lower() for value in [
                        transaction['person'].get('nombre_completo', ''),
                        transaction['person'].get('cedula', ''),
                        transaction['person'].get('cargo', ''),
                        transaction['person'].get('compania', ''),
                        transaction['descripcion'],
                        transaction['tipo_tarjeta'],
                        transaction['numero_tarjeta'],
                        transaction['fecha_transaccion'],
                        transaction['numero_autorizacion'],
                        transaction['categoria'],
                        transaction['subcategoria'],
                        transaction['zona'],
                        transaction['dia']
                    ]):
                        filtered_list.append(transaction)
                transactions_list = filtered_list
            # --- END FILTERING LOGIC ---
            
            print(f"DEBUG: Total transactions in transactions_list: {len(transactions_list)}")
            if transactions_list:
                print(f"DEBUG: First transaction in list: {transactions_list[0]}")
                print(f"DEBUG: Cedula for first transaction: {transactions_list[0].get('person', {}).get('cedula', 'N/A')}")
                print(f"DEBUG: Fecha_transaccion for first transaction: {transactions_list[0].get('fecha_transaccion', 'N/A')}")
            else:
                print("DEBUG: transactions_list is empty.")

            paginator = Paginator(transactions_list, 100)
            page_number = request.GET.get('page')
            page_obj = paginator.get_page(page_number)
            
            context['transactions'] = page_obj
            context['page_obj'] = page_obj
            context['paginator'] = paginator

        except FileNotFoundError:
            messages.error(request, "Error: tcs.xlsx not found in 'core/src/'. Please ensure the PDF processing has been run.")
            print(f"ERROR: tcs.xlsx not found at {tcs_excel_path}")
        except Exception as e:
            messages.error(request, f"Error reading or processing tcs.xlsx: {e}")
            print(f"CRITICAL ERROR: {e}")
    else:
        messages.warning(request, "tcs.xlsx not found. No transaction data to display.")
        print(f"WARNING: tcs.xlsx does not exist.")

    return render(request, 'tcs.html', context)

@login_required
def main(request):
    """
    Main dashboard view. Gathers counts for various data types and passes them to the home template.
    """
    context = {
        'person_count': Person.objects.count(),
        'conflict_count': Conflict.objects.count(),
        'finances_count': FinancialReport.objects.count(),
        'alerts_count': Person.objects.filter(revisar=True).count(),
        # Corrected count for Accionista del Grupo to count distinct persons
        'accionista_grupo_count': Person.objects.filter(conflicts__q3=True).distinct().count(), # Changed from conflict__q3 to conflicts__q3
        # Count for Aum. Pat. Subito > 2, as seen in the original home.html
        'aum_pat_subito_alert_count': FinancialReport.objects.filter(aum_pat_subito__gt=2).count(),
        # New counts for declarations per year
        'declarations_2021_count': FinancialReport.objects.filter(ano_declaracion=2021).count(),
        'declarations_2022_count': FinancialReport.objects.filter(ano_declaracion=2022).count(),
        'declarations_2023_count': FinancialReport.objects.filter(ano_declaracion=2023).count(),
        'declarations_2024_count': FinancialReport.objects.filter(ano_declaracion=2024).count(),
        # Corrected counts for conflicts per year, based on the 'fecha_inicio' field from the Conflict model
        'conflicts_2021_count': Conflict.objects.filter(fecha_inicio__year=2021).count(),
        'conflicts_2022_count': Conflict.objects.filter(fecha_inicio__year=2022).count(),
        'conflicts_2023_count': Conflict.objects.filter(fecha_inicio__year=2023).count(),
        'conflicts_2024_count': Conflict.objects.filter(fecha_inicio__year=2024).count(),
        # Count for active persons
        'active_person_count': Person.objects.filter(estado='Activo').count(),
        # Count for retired persons
        'retired_person_count': Person.objects.filter(estado='Retirado').count(),
        'tc_count': CreditCard.objects.count(),
    }

    return render(request, 'home.html', context)

@login_required
def import_persons(request):
    """View for importing persons data from Excel files"""
    if request.method == 'POST' and request.FILES.get('excel_file'):
        excel_file = request.FILES['excel_file']
        try:
            # Define the path to save the uploaded file temporarily
            temp_upload_path = os.path.join(settings.BASE_DIR, 'core', 'src', 'uploaded_persons_temp.xlsx')
            with open(temp_upload_path, 'wb+') as destination:
                for chunk in excel_file.chunks():
                    destination.write(chunk)

            # Read the Excel file into a pandas DataFrame
            df = pd.read_excel(temp_upload_path)

            # Remove the temporary uploaded file
            os.remove(temp_upload_path)

            # Strip whitespace and convert column names to lowercase for consistent mapping
            df.columns = df.columns.str.strip().str.lower()

            # Define column mapping from Excel columns to model fields
            column_mapping = {
                'id': 'id',
                'nombre completo': 'nombre_completo',
                'correo_normalizado': 'raw_correo',
                'cedula': 'cedula',
                'estado': 'estado',
                'compania': 'compania',
                'cargo': 'cargo',
                'activo': 'activo',
                'BUSINESS UNIT': 'area',
            }

            # Rename columns based on the mapping
            df = df.rename(columns=column_mapping)

            # Ensure 'estado' column exists, if 'activo' is present, use it to determine 'estado'
            if 'activo' in df.columns and 'estado' not in df.columns:
                df['estado'] = df['activo'].apply(lambda x: 'Activo' if x else 'Retirado')
            elif 'estado' not in df.columns:
                df['estado'] = 'Activo' # Default to 'Activo' if neither 'estado' nor 'activo' is present

            # Convert 'cedula' to string type to prevent issues with mixed types
            if 'cedula' in df.columns:
                df['cedula'] = df['cedula'].astype(str)
            else:
                messages.error(request, "Error: 'Cedula' column not found in the Excel file.")
                return HttpResponseRedirect('/import/')

            # Convert nombre_completo to title case if it exists
            if 'nombre_completo' in df.columns:
                df['nombre_completo'] = df['nombre_completo'].str.title()

            # Process 'raw_correo' to create 'correo_to_use' for the database and output
            if 'raw_correo' in df.columns:
                df['correo_to_use'] = df['raw_correo'].str.lower()
            else:
                df['correo_to_use'] = '' # Initialize if no raw email is present

            # Define the columns for the output Excel file including 'Id', 'Estado', and the new 'correo' and 'AREA'
            output_columns = ['Id', 'NOMBRE COMPLETO', 'Cedula', 'Estado', 'Compania', 'CARGO', 'correo', 'AREA']
            output_columns_df = pd.DataFrame(columns=output_columns)

            # Populate the output DataFrame with data from the processed DataFrame
            if 'id' in df.columns:
                output_columns_df['Id'] = df['id']
            if 'nombre_completo' in df.columns:
                output_columns_df['NOMBRE COMPLETO'] = df['nombre_completo']
            if 'cedula' in df.columns:
                output_columns_df['Cedula'] = df['cedula']
            if 'estado' in df.columns:
                output_columns_df['Estado'] = df['estado']
            if 'compania' in df.columns:
                output_columns_df['Compania'] = df['compania']
            if 'cargo' in df.columns:
                output_columns_df['CARGO'] = df['cargo']
            if 'correo_to_use' in df.columns:
                output_columns_df['correo'] = df['correo_to_use']
            # Add 'AREA' to the output DataFrame
            if 'area' in df.columns:
                output_columns_df['AREA'] = df['area']
            else:
                output_columns_df['AREA'] = '' # Ensure column exists even if no data

            # Define the path for the output Excel file
            output_excel_path = os.path.join(settings.BASE_DIR, 'core', 'src', 'Personas.xlsx')

            # Save the filtered and formatted DataFrame to a new Excel file
            output_columns_df.to_excel(output_excel_path, index=False)

            # Iterate over the DataFrame and update/create Person objects in the database
            for _, row in df.iterrows():
                Person.objects.update_or_create(
                    cedula=row['cedula'],
                    defaults={
                        'nombre_completo': row.get('nombre_completo', ''),
                        'correo': row.get('correo_to_use', ''),
                        'estado': row.get('estado', 'Activo'),
                        'compania': row.get('compania', ''),
                        'cargo': row.get('cargo', ''),
                        'area': row.get('area', ''), 
                    }
                )

            messages.success(request, f'Archivo de personas importado exitosamente! {len(df)} registros procesados y Personas.xlsx generado.')
        except Exception as e:
            messages.error(request, f'Error procesando archivo de personas: {str(e)}')

        return HttpResponseRedirect('/import/')

    return HttpResponseRedirect('/import/')

@login_required
def import_conflicts(request):
    """View for importing conflicts data from Excel files"""
    if request.method == 'POST' and request.FILES.get('conflict_excel_file'):
        excel_file = request.FILES['conflict_excel_file']
        try:
            dest_path = os.path.join(settings.BASE_DIR, "core", "src", "conflictos.xlsx")
            with open(dest_path, 'wb+') as destination:
                for chunk in excel_file.chunks():
                    destination.write(chunk)

            subprocess.run(['python3', 'core/conflicts.py'], check=True, cwd=settings.BASE_DIR)

            processed_file = os.path.join(settings.BASE_DIR, "core", "src", "conflicts.xlsx")
            df = pd.read_excel(processed_file)
            df.columns = df.columns.str.lower().str.replace(' ', '_')

            # Helper function to process boolean fields
            def get_boolean_value(value):
                if pd.isna(value):
                    return None  # Return None for NaN/empty values
                return bool(value) # Convert to boolean otherwise

            for _, row in df.iterrows():
                try:
                    person, created = Person.objects.get_or_create(
                        cedula=str(row['cedula']),
                        defaults={
                            'nombre_completo': row.get('nombre', ''),
                            'correo': row.get('email', ''),
                            'compania': row.get('compañía', ''),
                            'cargo': row.get('cargo', '')
                        }
                    )

                    fecha_inicio_str = row.get('fecha_de_inicio')
                    fecha_inicio_date = None
                    if pd.notna(fecha_inicio_str):
                        try:
                            fecha_inicio_date = pd.to_datetime(fecha_inicio_str).date()
                        except ValueError:
                            messages.warning(request, f"Could not parse date '{fecha_inicio_str}' for conflict. Skipping row.")
                            continue

                    Conflict.objects.update_or_create(
                        person=person,
                        fecha_inicio=fecha_inicio_date,
                        defaults={
                            'q1': get_boolean_value(row.get('q1')), # Changed
                            'q1_detalle': row.get('q1_detalle', ''),
                            'q2': get_boolean_value(row.get('q2')), # Changed
                            'q2_detalle': row.get('q2_detalle', ''),
                            'q3': get_boolean_value(row.get('q3')), # Changed
                            'q3_detalle': row.get('q3_detalle', ''),
                            'q4': get_boolean_value(row.get('q4')), # Changed
                            'q4_detalle': row.get('q4_detalle', ''),
                            'q5': get_boolean_value(row.get('q5')), # Changed
                            'q5_detalle': row.get('q5_detalle', ''),
                            'q6': get_boolean_value(row.get('q6')), # Changed
                            'q6_detalle': row.get('q6_detalle', ''),
                            'q7': get_boolean_value(row.get('q7')), # Changed
                            'q7_detalle': row.get('q7_detalle', ''),
                            'q8': get_boolean_value(row.get('q8')), # Changed
                            'q9': get_boolean_value(row.get('q9')), # Changed
                            'q10': get_boolean_value(row.get('q10')), # Changed
                            'q10_detalle': row.get('q10_detalle', ''),
                            'q11': get_boolean_value(row.get('q11')), # Changed
                            'q11_detalle': row.get('q11_detalle', '')
                        }
                    )

                except Exception as e:
                    messages.error(request, f"Error processing row with cedula {row.get('cedula', 'N/A')}: {str(e)}")
                    continue

            messages.success(request, f'Archivo de conflictos importado exitosamente! {len(df)} registros procesados.')
        except Exception as e:
            messages.error(request, f'Error procesando archivo de conflictos: {str(e)}')

        return HttpResponseRedirect('/import/')

    return HttpResponseRedirect('/import/')

@login_required
def import_financial_reports(request):
    """View for importing financial reports data from idTrends.xlsx"""
    # This function is called internally after idTrends.py generates the file
    # It does not expect a file upload directly from the user.
    try:
        file_path = os.path.join(settings.BASE_DIR, 'core', 'src', 'idTrends.xlsx')
        if not os.path.exists(file_path):
            messages.error(request, "Error: idTrends.xlsx not found. Please ensure analysis scripts run first.")
            return

        df = pd.read_excel(file_path)
        # Ensure column names are consistently lowercased and spaces/dots are replaced
        df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_').str.replace('.', '', regex=False).str.replace('á', 'a').str.replace('é', 'e').str.replace('í', 'i').str.replace('ó', 'o').str.replace('ú', 'u')

        # No need for column_mapping dictionary if direct access is used after cleaning column names
        for _, row in df.iterrows():
            try:
                cedula = str(row.get('cedula'))
                if not cedula:
                    messages.warning(request, f"Skipping row due to missing cedula: {row.to_dict()}")
                    continue # Skip rows without a cedula

                person, created = Person.objects.get_or_create(
                    cedula=cedula,
                    defaults={
                        'nombre_completo': row.get('nombre_completo', ''),
                        'correo': row.get('correo', ''),
                        'compania': row.get('compania_y', ''), # Use compania_y from idTrends.py output
                        'cargo': row.get('cargo', '')
                    }
                )

                apalancamiento_val = _clean_numeric_value(row.get('apalancamiento'))
                # If apalancamiento_val is a number > 1 and it didn't originally have a '%' sign,
                # it's likely a percentage like 12.45 which needs to be stored as 0.1245
                if isinstance(apalancamiento_val, (int, float)) and apalancamiento_val is not None and apalancamiento_val > 1.0 and '%' not in str(row.get('apalancamiento', '')):
                    apalancamiento_val /= 100

                endeudamiento_val = _clean_numeric_value(row.get('endeudamiento'))
                # Similar logic for endeudamiento_val
                if isinstance(endeudamiento_val, (int, float)) and endeudamiento_val is not None and endeudamiento_val > 1.0 and '%' not in str(row.get('endeudamiento', '')):
                    endeudamiento_val /= 100

                # Prepare data for FinancialReport, handling potential NaN values and cleaning numeric fields
                report_data = {
                    'person': person,
                    'fk_id_periodo': _clean_numeric_value(row.get('fkidperiodo')), # Corrected column name access
                    'ano_declaracion': _clean_numeric_value(row.get('año_declaracion')),
                    'ano_creacion': _clean_numeric_value(row.get('año_creacion')),
                    'activos': _clean_numeric_value(row.get('activos')),
                    'cant_bienes': _clean_numeric_value(row.get('cant_bienes')),
                    'cant_bancos': _clean_numeric_value(row.get('cant_bancos')),
                    'cant_cuentas': _clean_numeric_value(row.get('cant_cuentas')),
                    'cant_inversiones': _clean_numeric_value(row.get('cant_inversiones')),
                    'pasivos': _clean_numeric_value(row.get('pasivos')),
                    'cant_deudas': _clean_numeric_value(row.get('cant_deudas')),
                    'patrimonio': _clean_numeric_value(row.get('patrimonio')),
                    'apalancamiento': apalancamiento_val, # Use the processed value
                    'endeudamiento': endeudamiento_val,
                    'capital': _clean_numeric_value(row.get('capital')),
                    'aum_pat_subito': _clean_numeric_value(row.get('aum_pat_subito')), # Apply cleaning here
                    'activos_var_abs': _clean_numeric_value(row.get('activos_var_abs')),
                    'activos_var_rel': str(row.get('activos_var_rel')) if pd.notna(row.get('activos_var_rel')) else '',
                    'pasivos_var_abs': _clean_numeric_value(row.get('pasivos_var_abs')),
                    'pasivos_var_rel': str(row.get('pasivos_var_rel')) if pd.notna(row.get('pasivos_var_rel')) else '',
                    'patrimonio_var_abs': _clean_numeric_value(row.get('patrimonio_var_abs')),
                    'patrimonio_var_rel': str(row.get('patrimonio_var_rel')) if pd.notna(row.get('patrimonio_var_rel')) else '',
                    'apalancamiento_var_abs': _clean_numeric_value(row.get('apalancamiento_var_abs')),
                    'apalancamiento_var_rel': str(row.get('apalancamiento_var_rel')) if pd.notna(row.get('apalancamiento_var_rel')) else '',
                    'endeudamiento_var_abs': _clean_numeric_value(row.get('endeudamiento_var_abs')),
                    'endeudamiento_var_rel': str(row.get('endeudamiento_var_rel')) if pd.notna(row.get('endeudamiento_var_rel')) else '',
                    'banco_saldo': _clean_numeric_value(row.get('banco_saldo')),
                    'bienes': _clean_numeric_value(row.get('bienes')),
                    'inversiones': _clean_numeric_value(row.get('inversiones')),
                    'banco_saldo_var_abs': _clean_numeric_value(row.get('banco_saldo_var_abs')),
                    'banco_saldo_var_rel': str(row.get('banco_saldo_var_rel')) if pd.notna(row.get('banco_saldo_var_rel')) else '',
                    'bienes_var_abs': _clean_numeric_value(row.get('bienes_var_abs')),
                    'bienes_var_rel': str(row.get('bienes_var_rel')) if pd.notna(row.get('bienes_var_rel')) else '',
                    'inversiones_var_abs': _clean_numeric_value(row.get('inversiones_var_abs')),
                    'inversiones_var_rel': str(row.get('inversiones_var_rel')) if pd.notna(row.get('inversiones_var_rel')) else '',
                    'ingresos': _clean_numeric_value(row.get('ingresos')),
                    'cant_ingresos': _clean_numeric_value(row.get('cant_ingresos')),
                    'ingresos_var_abs': _clean_numeric_value(row.get('ingresos_var_abs')),
                    'ingresos_var_rel': str(row.get('ingresos_var_rel')) if pd.notna(row.get('ingresos_var_rel')) else '',
                }

                # Use update_or_create based on person and fk_id_periodo to ensure uniqueness per period
                # Ensure fk_id_periodo is not None for update_or_create to work correctly
                if report_data['fk_id_periodo'] is not None:
                    FinancialReport.objects.update_or_create(
                        person=person,
                        fk_id_periodo=report_data['fk_id_periodo'],
                        defaults=report_data
                    )
                else:
                    messages.warning(request, f"Skipping row for {person.nombre_completo} due to missing fk_id_periodo.")

            except Exception as e:
                messages.error(request, f"Error processing financial report for row: {row.to_dict()}. Error: {e}")
                
        messages.success(request, f"Se importaron exitosamente los datos de {len(df)} reportes financieros.")

    except FileNotFoundError:
        messages.error(request, "Error: idTrends.xlsx no se encontró. Asegúrese de que los scripts de análisis se hayan ejecutado correctamente.")
    except Exception as e:
        messages.error(request, f"Ocurrió un error al procesar el archivo idTrends.xlsx: {e}")

    return redirect('import')

@login_required
def person_list(request):
    search_query = request.GET.get('q', '')
    status_filter = request.GET.get('status', '')
    cargo_filter = request.GET.get('cargo', '')
    compania_filter = request.GET.get('compania', '')
    area_filter = request.GET.get('area', '')

    order_by = request.GET.get('order_by', 'nombre_completo')
    sort_direction = request.GET.get('sort_direction', 'asc')

    persons = Person.objects.all()

    if search_query:
        persons = persons.filter(
            Q(nombre_completo__icontains=search_query) |
            Q(cedula__icontains=search_query) |
            Q(correo__icontains=search_query) |
            Q(area__icontains=search_query) 
        )

    if status_filter:
        persons = persons.filter(estado=status_filter)

    if cargo_filter:
        persons = persons.filter(cargo=cargo_filter)

    if compania_filter:
        persons = persons.filter(compania=compania_filter)

    if area_filter: # New: Apply area filter
        persons = persons.filter(area=area_filter)

    if sort_direction == 'desc':
        order_by = f'-{order_by}'
    persons = persons.order_by(order_by)

    # Convert names to title case for display
    for person in persons:
        person.nombre_completo = person.nombre_completo.title()

    cargos = Person.objects.exclude(cargo='').values_list('cargo', flat=True).distinct().order_by('cargo')
    companias = Person.objects.exclude(compania='').values_list('compania', flat=True).distinct().order_by('compania')
    areas = Person.objects.exclude(area='').values_list('area', flat=True).distinct().order_by('area') 

    paginator = Paginator(persons, 25)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    context = {
        'persons': page_obj,
        'page_obj': page_obj,
        'cargos': cargos,
        'companias': companias,
        'areas': areas,
        'current_order': order_by.lstrip('-'),
        'current_direction': 'desc' if order_by.startswith('-') else 'asc',
        'all_params': {k: v for k, v in request.GET.items() if k not in ['page', 'order_by', 'sort_direction']},
        'alerts_count': Person.objects.filter(revisar=True).count(), # Add alerts count
    }

    return render(request, 'persons.html', context)

@login_required
def export_persons_excel(request):
    search_query = request.GET.get('q', '')
    status_filter = request.GET.get('status', '')
    cargo_filter = request.GET.get('cargo', '')
    compania_filter = request.GET.get('compania', '')
    revisar_filter = request.GET.get('revisar', '') # <--- Add this line to get the 'revisar' parameter

    order_by = request.GET.get('order_by', 'nombre_completo')
    sort_direction = request.GET.get('sort_direction', 'asc')

    persons = Person.objects.all()

    # Apply the 'revisar' filter if present in the URL
    if revisar_filter == 'True': # <--- Add this block
        persons = persons.filter(revisar=True)

    if search_query:
        persons = persons.filter(
            Q(nombre_completo__icontains=search_query) |
            Q(cedula__icontains=search_query) |
            Q(correo__icontains=search_query)
        )

    if status_filter:
        persons = persons.filter(estado=status_filter)

    if cargo_filter:
        persons = persons.filter(cargo=cargo_filter)

    if compania_filter:
        persons = persons.filter(compania=compania_filter)

    # --- Add dynamic column filtering for FinancialReport fields ---
    i = 0
    while f'column_{i}' in request.GET:
        column = request.GET.get(f'column_{i}')
        operator = request.GET.get(f'operator_{i}')
        value1 = request.GET.get(f'value_{i}')
        value2 = request.GET.get(f'value2_{i}')

        if column and operator and value1:
            # Corrected: Use 'financial_reports' as the related name from Person to FinancialReport
            filter_key = f'financial_reports__{column}'

            try:
                # Remove commas from value1 and value2 before conversion
                if isinstance(value1, str):
                    value1 = value1.replace(',', '')
                if isinstance(value2, str):
                    value2 = value2.replace(',', '')

                # Convert value1 to appropriate type based on common financial fields
                if column in ['fk_id_periodo', 'ano_declaracion', 'cant_bienes', 'cant_bancos', 'cant_cuentas',
                              'cant_inversiones', 'cant_deudas', 'cant_ingresos']:
                    value1 = int(float(value1)) # Convert to int if it's a count/ID
                    if value2: value2 = int(float(value2))
                else: # Assume float for monetary values and percentages
                    value1 = float(value1)
                    if value2: value2 = float(value2)
            except (ValueError, TypeError):
                # Handle cases where conversion fails (e.g., non-numeric input for numeric fields)
                # You might want to log this or provide user feedback
                value1 = None # Invalidate the filter if value is not convertible
                value2 = None

            if value1 is not None:
                if operator == '>':
                    persons = persons.filter(**{f'{filter_key}__gt': value1})
                elif operator == '<':
                    persons = persons.filter(**{f'{filter_key}__lt': value1})
                elif operator == '=':
                    persons = persons.filter(**{f'{filter_key}': value1})
                elif operator == '>=':
                    persons = persons.filter(**{f'{filter_key}__gte': value1})
                elif operator == '<=':
                    persons = persons.filter(**{f'{filter_key}__lte': value1})
                elif operator == 'between' and value2 is not None:
                    persons = persons.filter(**{f'{filter_key}__range': (min(value1, value2), max(value1, value2))})
                elif operator == 'contains':
                    # 'contains' operator is typically for text fields.
                    # Ensure the column is a text field or handle accordingly.
                    # For numeric fields, 'contains' usually doesn't make sense.
                    persons = persons.filter(**{f'{filter_key}__icontains': str(value1)})
        i += 1


    if sort_direction == 'desc':
        order_by = f'-{order_by}'
    persons = persons.order_by(order_by).distinct() # Use .distinct() to avoid duplicate persons if related objects cause issues

    # Prepare data for DataFrame
    data = []
    for person in persons:
        # You might need to adjust this logic based on how you want to handle multiple reports per person
        financial_report = FinancialReport.objects.filter(person=person).order_by('-ano_declaracion', '-fk_id_periodo').first()

        row_data = {
            'ID': person.cedula,
            'Nombre Completo': person.nombre_completo,
            'Correo': person.correo,
            'Estado': person.estado,
            'Compañía': person.compania,
            'Cargo': person.cargo,
            'Revisar': 'Sí' if person.revisar else 'No',
            'Comentarios': person.comments,
            'Creado En': person.created_at.strftime('%Y-%m-%d %H:%M:%S') if person.created_at else '',
            'Actualizado En': person.updated_at.strftime('%Y-%m-%d %H:%M:%S') if person.updated_at else '',
        }

        # Add financial report data if available
        if financial_report:
            row_data.update({
                'Periodo': financial_report.fk_id_periodo,
                'Ano': financial_report.ano_declaracion,
                'Aum. Pat. Subito': financial_report.aum_pat_subito,
                '% Endeudamiento': financial_report.endeudamiento,
                'Patrimonio': financial_report.patrimonio,
                'Patrimonio Var. Rel. %': financial_report.patrimonio_var_rel,
                'Patrimonio Var. Abs. $': financial_report.patrimonio_var_abs,
                'Activos': financial_report.activos,
                'Activos Var. Rel. %': financial_report.activos_var_rel,
                'Activos Var. Abs. $': financial_report.activos_var_abs,
                'Pasivos': financial_report.pasivos,
                'Pasivos Var. Rel. %': financial_report.pasivos_var_rel,
                'Pasivos Var. Abs. $': financial_report.pasivos_var_abs,
                'Cant. Deudas': financial_report.cant_deudas,
                'Ingresos': financial_report.ingresos,
                'Ingresos Var. Rel. %': financial_report.ingresos_var_rel,
                'Ingresos Var. Abs. $': financial_report.ingresos_var_abs,
                'Cant. Ingresos': financial_report.cant_ingresos,
                'Bancos Saldo': financial_report.banco_saldo,
                'Bancos Var. Rel. %': financial_report.banco_saldo_var_rel,
                'Bancos Var. $': financial_report.banco_saldo_var_abs,
                'Cant. Cuentas': financial_report.cant_cuentas,
                'Cant. Bancos': financial_report.cant_bancos,
                'Bienes Valor': financial_report.bienes,
                'Bienes Var. Rel. %': financial_report.bienes_var_rel,
                'Bienes Var. $': financial_report.bienes_var_abs,
                'Cant. Bienes': financial_report.cant_bienes,
                'Inversiones Valor': financial_report.inversiones,
                'Inversiones Var. Rel. %': financial_report.inversiones_var_rel,
                'Inversiones Var. $': financial_report.inversiones_var_abs,
                'Cant. Inversiones': financial_report.cant_inversiones,
            })
        data.append(row_data)

    df = pd.DataFrame(data)

    # Create an in-memory Excel file
    excel_file = io.BytesIO()
    with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Persons', index=False)
    excel_file.seek(0)

    # Create the HTTP response
    response = HttpResponse(excel_file.read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = f'attachment; filename="persons_export_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx"'
    return response

@login_required
def conflict_list(request):
    search_query = request.GET.get('q', '')
    compania_filter = request.GET.get('compania', '')
    column_filter = request.GET.get('column', '')
    answer_filter = request.GET.get('answer', '')
    missing_details_view = request.GET.get('missing_details', False) # New parameter for missing details

    order_by = request.GET.get('order_by', 'person__nombre_completo')
    sort_direction = request.GET.get('sort_direction', 'asc')

    conflicts = Conflict.objects.select_related('person').all()

    if search_query:
        conflicts = conflicts.filter(
            Q(person__nombre_completo__icontains=search_query) |
            Q(person__cedula__icontains=search_query)
        )

    if compania_filter:
        conflicts = conflicts.filter(person__compania=compania_filter)

    if column_filter and answer_filter:
        filter_q = Q()
        if answer_filter == 'yes':
            filter_q = Q(**{column_filter: True})
        elif answer_filter == 'no':
            filter_q = Q(**{column_filter: False})
        elif answer_filter == 'blank': # Filter for blank answers
            filter_q = Q(**{column_filter: None})
        conflicts = conflicts.filter(filter_q)

    # Filtering for conflicts where qX is True but qX_detalle is blank (None or empty string)
    if missing_details_view == 'true': # Check if the parameter is explicitly 'true'
        missing_details_q = Q()
        for i in range(1, 12):
            q_field = f'q{i}'
            detail_field = f'q{i}_detalle'
            # Only apply this if the q_field is True AND the detail_field is either None or an empty string
            missing_details_q |= Q(**{q_field: True, detail_field + '__isnull': True}) | Q(**{q_field: True, detail_field: ''})
        conflicts = conflicts.filter(missing_details_q)

    if sort_direction == 'desc':
        order_by = f'-{order_by}'
    conflicts = conflicts.order_by(order_by)

    # Attach the '_detalle' fields as '_answer' for template display
    for conflict in conflicts:
        for i in range(1, 12):
            detail_field_name = f'q{i}_detalle'
            answer_field_name = f'q{i}_answer'
            # Use getattr to safely access the attribute, with a default of None if it doesn't exist
            setattr(conflict, answer_field_name, getattr(conflict, detail_field_name, None))


    companias = Person.objects.exclude(compania='').values_list('compania', flat=True).distinct().order_by('compania')

    paginator = Paginator(conflicts, 25)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    # Build all_params for pagination links
    all_params = {k: v for k, v in request.GET.items() if k not in ['page', 'order_by', 'sort_direction']}

    context = {
        'conflicts': page_obj,
        'page_obj': page_obj,
        'companias': companias,
        'current_order': order_by.lstrip('-'),
        'current_direction': 'desc' if order_by.startswith('-') else 'asc',
        'all_params': all_params,
        'alerts_count': Person.objects.filter(revisar=True).count(), # Add alerts count
        'missing_details_view': missing_details_view, # Pass the boolean to the template
    }
    return render(request, 'conflicts.html', context)

@login_required
def financial_report_list(request):
    # Initialize a Q object to accumulate all filters
    all_filters_q = Q()

    # Handle the main search query 'q' (for person's name or cedula)
    search_query = request.GET.get('q', '')
    if search_query:
        all_filters_q &= (
            Q(person__nombre_completo__icontains=search_query) |
            Q(person__cedula__icontains=search_query)
        )

    # Handle existing 'compania' filter
    compania_filter = request.GET.get('compania', '')
    if compania_filter:
        all_filters_q &= Q(person__compania=compania_filter)

    # Handle existing 'ano_declaracion' filter
    ano_declaracion_filter = request.GET.get('ano_declaracion', '')
    if ano_declaracion_filter:
        try:
            # Ensure it's an integer for exact match
            ano_declaracion_int = int(ano_declaracion_filter)
            all_filters_q &= Q(ano_declaracion=ano_declaracion_int)
        except ValueError:
            messages.warning(request, "Año de declaración inválido.")

    # Iterate through potential filter indices (e.g., column_0, column_1, etc.)
    i = 0
    while True:
        # The names in the GET request will be like column_0, operator_0, value_0, value2_0
        column = request.GET.get(f'column_{i}')
        operator = request.GET.get(f'operator_{i}')
        value1 = request.GET.get(f'value_{i}')
        value2 = request.GET.get(f'value2_{i}') # For 'between' operator

        # If no column is found for the current index, stop iterating
        if not column:
            break

        # Only apply a filter if column, operator, and at least value1 are present
        if column and operator and value1:
            try:
                if operator == '>':
                    all_filters_q &= Q(**{f"{column}__gt": _clean_numeric_value(value1)})
                elif operator == '<':
                    all_filters_q &= Q(**{f"{column}__lt": _clean_numeric_value(value1)})
                elif operator == '=':
                    # For exact match, for numeric fields use _clean_numeric_value
                    # For text fields, __iexact is often better than just =
                    cleaned_value = _clean_numeric_value(value1)
                    if cleaned_value is not None: # It's a number
                        all_filters_q &= Q(**{f"{column}": cleaned_value})
                    else: # Treat as text
                        all_filters_q &= Q(**{f"{column}__iexact": value1})
                elif operator == '>=':
                    all_filters_q &= Q(**{f"{column}__gte": _clean_numeric_value(value1)})
                elif operator == '<=':
                    all_filters_q &= Q(**{f"{column}__lte": _clean_numeric_value(value1)})
                elif operator == 'between' and value2:
                    val1_cleaned = _clean_numeric_value(value1)
                    val2_cleaned = _clean_numeric_value(value2)
                    if val1_cleaned is not None and val2_cleaned is not None:
                        # Ensure min/max for correct range
                        all_filters_q &= Q(**{f"{column}__range": (min(val1_cleaned, val2_cleaned), max(val1_cleaned, val2_cleaned))})
                    else:
                        messages.warning(request, f"Valores inválidos para el filtro 'entre' en columna {column}.")
                elif operator == 'contains':
                    # 'contains' is typically for text fields. Use icontains for case-insensitivity.
                    all_filters_q &= Q(**{f"{column}__icontains": value1})
                else:
                    messages.warning(request, f"Operador inválido '{operator}' para la columna {column}.")
            except ValueError:
                messages.error(request, f"Error al convertir valor para el filtro en {column}. Verifique el formato numérico.")
            except Exception as e:
                messages.error(request, f"Error inesperado al aplicar filtro en {column}: {e}")
        
        i += 1 # Move to the next potential filter index

    # Apply all accumulated filters to the queryset
    financial_reports = FinancialReport.objects.select_related('person').filter(all_filters_q)

    # Ordering logic
    order_by = request.GET.get('order_by', 'person__nombre_completo')
    sort_direction = request.GET.get('sort_direction', 'asc')

    if sort_direction == 'desc':
        order_by = f'-{order_by}'
    
    financial_reports = financial_reports.order_by(order_by)

    # Get distinct values for existing filters (Companias and Anos Declaracion)
    companias = Person.objects.exclude(compania='').values_list('compania', flat=True).distinct().order_by('compania')
    anos_declaracion = FinancialReport.objects.exclude(ano_declaracion__isnull=True).values_list('ano_declaracion', flat=True).distinct().order_by('ano_declaracion')

    # Pagination
    paginator = Paginator(financial_reports, 25)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    # Prepare all_params for pagination links to persist all GET parameters
    all_params = request.GET.copy()
    if 'page' in all_params:
        del all_params['page']
    
    # Alert count
    alerts_count = Person.objects.filter(revisar=True).count()

    context = {
        'financial_reports': page_obj,
        'page_obj': page_obj,
        'companias': companias,
        'anos_declaracion': anos_declaracion,
        'current_order': order_by.lstrip('-'),
        'current_direction': 'desc' if order_by.startswith('-') else 'asc',
        'all_params': all_params, # This will now correctly include all column_X, operator_X, value_X params
        'alerts_count': alerts_count,
    }

    return render(request, 'finances.html', context)

@login_required
def import_finances(request):
    """View for importing protected Excel files and running analysis"""
    if request.method == 'POST' and request.FILES.get('finances_file'):
        excel_file = request.FILES['finances_file']
        password = request.POST.get('excel_password', '')

        try:
            # Save the original file temporarily
            temp_path = os.path.join(settings.BASE_DIR, 'core', 'src', 'temp_protected.xlsx')
            with open(temp_path, 'wb+') as destination:
                for chunk in excel_file.chunks():
                    destination.write(chunk)

            # Try to decrypt the file if password is provided
            decrypted_path = os.path.join(settings.BASE_DIR, 'core', 'src', 'data.xlsx')

            if password:
                try:
                    with open(temp_path, 'rb') as f:
                        file = msoffcrypto.OfficeFile(f)
                        file.load_key(password=password)
                        decrypted = io.BytesIO()
                        file.decrypt(decrypted)

                        with open(decrypted_path, 'wb') as out:
                            out.write(decrypted.getvalue())
                except Exception as e:
                    messages.error(request, f'Error al desproteger el archivo: {str(e)}')
                    return HttpResponseRedirect('/import/')
            else:
                # If no password, just copy the file
                import shutil
                shutil.copyfile(temp_path, decrypted_path)

            # Remove the temporary file
            os.remove(temp_path)

            # Run the analysis scripts in sequence
            try:
                # Run cats.py analysis
                subprocess.run(['python3', 'core/cats.py'], check=True, cwd=settings.BASE_DIR)

                # Run nets.py analysis
                subprocess.run(['python3', 'core/nets.py'], check=True, cwd=settings.BASE_DIR)

                # Run trends.py analysis
                subprocess.run(['python3', 'core/trends.py'], check=True, cwd=settings.BASE_DIR)

                # Run idTrends.py analysis
                subprocess.run(['python3', 'core/idTrends.py'], check=True, cwd=settings.BASE_DIR)

                # After idTrends.py generates idTrends.xlsx, import the data into the FinancialReport model
                import_financial_reports(request) # Call the new import function

                # Remove the data.xlsx file after processing
                os.remove(decrypted_path)

                messages.success(request, 'Archivo procesado exitosamente y análisis completado!')
            except subprocess.CalledProcessError as e:
                messages.error(request, f'Error ejecutando análisis: {str(e)}')
            except Exception as e:
                messages.error(request, f'Error durante el análisis: {str(e)}')

        except Exception as e:
            messages.error(request, f'Error procesando archivo protegido: {str(e)}')

        return HttpResponseRedirect('/import/')

    return HttpResponseRedirect('/import/')

@login_required
def person_details(request, cedula):
    """
    View to display the details of a specific person, including related financial reports, conflicts, and credit card transactions.
    """
    myperson = get_object_or_404(Person, cedula=cedula)
    financial_reports = FinancialReport.objects.filter(person=myperson).order_by('-ano_declaracion')
    conflicts = myperson.conflicts.all().order_by('-fecha_inicio')
    
    # Retrieve all CreditCard objects associated with the person
    credit_card_transactions = CreditCard.objects.filter(person=myperson)
    
    context = {
        'myperson': myperson,
        'financial_reports': financial_reports,
        'conflicts': conflicts,
        'alerts_count': Person.objects.filter(revisar=True).count(),
        'credit_card_transactions': credit_card_transactions, # Add the credit card transactions to the context
    }
    
    return render(request, 'details.html', context)


@login_required
def alerts_list(request):
    """
    View to display persons marked for review (revisar=True).
    """
    search_query = request.GET.get('q', '')
    order_by = request.GET.get('order_by', 'nombre_completo')
    sort_direction = request.GET.get('sort_direction', 'asc')

    persons = Person.objects.filter(revisar=True)

    if search_query:
        persons = persons.filter(
            Q(nombre_completo__icontains=search_query) |
            Q(cedula__icontains=search_query) |
            Q(correo__icontains=search_query))

    if sort_direction == 'desc':
        order_by = f'-{order_by}'
    persons = persons.order_by(order_by)

    # Convert names to title case for display
    for person in persons:
        person.nombre_completo = person.nombre_completo.title()

    paginator = Paginator(persons, 25)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    context = {
        'persons': page_obj,
        'page_obj': page_obj,
        'current_order': order_by.lstrip('-'),
        'current_direction': 'desc' if order_by.startswith('-') else 'asc',
        'all_params': {k: v for k, v in request.GET.items() if k not in ['page', 'order_by', 'sort_direction']},
        'alerts_count': Person.objects.filter(revisar=True).count(), # Pass alerts count
    }

    return render(request, 'alerts.html', context)

@require_POST
def save_comment(request, cedula):
    person = get_object_or_404(Person, cedula=cedula)
    new_comment = request.POST.get('new_comment')

    if new_comment:
        now = datetime.now() 
        timestamp = now.strftime("%d/%m/%Y %H:%M:%S")

        formatted_comment = f"[{timestamp}] {new_comment}"
        
        if person.comments:
            person.comments += f"\n{formatted_comment}"
        else:
            person.comments = formatted_comment
        
        person.save()

    return redirect('person_details', cedula=cedula)

@require_POST
def delete_comment(request, cedula, comment_index):
    person = get_object_or_404(Person, cedula=cedula)

    if person.comments:
        comments_list = person.comments.splitlines()
        
        # Filter out empty strings that might result from splitlines
        comments_list = [comment.strip() for comment in comments_list if comment.strip()]

        if 0 <= comment_index < len(comments_list):
            # Remove the comment at the specified index
            comments_list.pop(comment_index)
            
            # Join the remaining comments back into a single string
            person.comments = "\n".join(comments_list)
            person.save()
        else:
            # Handle invalid index, e.g., log an error or show a message
            pass # For now, silently ignore invalid index
    
    return redirect('person_details', cedula=cedula)

# Updated import_tcs function
def import_tcs(request):
    if request.method == 'POST':
        pdf_files = request.FILES.getlist('visa_pdf_files')
        pdf_password = request.POST.get('visa_pdf_password', '')

        if not pdf_files:
            messages.error(request, 'No se seleccionaron archivos PDF.', extra_tags='import_tcs')
            return redirect('import_page')

        input_pdf_dir = os.path.join(settings.BASE_DIR, 'core', 'src', 'extractos')
        output_excel_dir = os.path.join(settings.BASE_DIR, 'core', 'src')
        tcs_excel_path = os.path.join(output_excel_dir, "tcs.xlsx") # Path to the output Excel

        os.makedirs(input_pdf_dir, exist_ok=True) # Ensure input directory exists

        # Clear existing PDFs in the input_pdf_dir before saving new ones
        for filename in os.listdir(input_pdf_dir):
            if filename.endswith(".pdf"):
                os.remove(os.path.join(input_pdf_dir, filename))

        files_saved = 0
        for pdf_file in pdf_files:
            file_path = os.path.join(input_pdf_dir, pdf_file.name)
            try:
                with open(file_path, 'wb+') as destination:
                    for chunk in pdf_file.chunks():
                        destination.write(chunk)
                files_saved += 1
            except Exception as e:
                messages.error(request, f"Error saving PDF '{pdf_file.name}': {e}", extra_tags='import_tcs')

        if files_saved > 0:
            try:
                tcs.pdf_password = pdf_password
                tcs.run_pdf_processing(settings.BASE_DIR, input_pdf_dir, output_excel_dir)

                # --- NEW LOGIC: Load tcs.xlsx into CreditCard model ---
                if os.path.exists(tcs_excel_path):
                    # Force 'Cedula' column to be read as string to prevent float interpretation
                    df_tcs = pd.read_excel(tcs_excel_path, dtype={'Cedula': str})
                    transactions_created = 0
                    transactions_updated = 0

                    # --- NEW: Define clean_cedula_format locally for views.py ---
                    def clean_cedula_format(value):
                        try:
                            if isinstance(value, float) and value.is_integer():
                                return str(int(value))
                            return str(value)
                        except (ValueError, TypeError):
                            return str(value)

                    for index, row in df_tcs.iterrows():
                        raw_cedula = row.get('Cedula') # Get Cedula from the DataFrame row
                        cleaned_cedula = clean_cedula_format(raw_cedula) # Apply cleaning here

                        if pd.isna(cleaned_cedula) or cleaned_cedula == '':
                            print(f"Skipping row {index}: Missing or invalid Cedula for transaction {row.get('Descripción')}")
                            continue

                        # Find or create the Person.
                        person_obj, created = Person.objects.get_or_create(
                            cedula=cleaned_cedula, # Use the cleaned cedula
                            defaults={
                                'nombre_completo': row.get('Tarjetahabiente', ''),
                                'cargo': row.get('CARGO', ''),
                                'compania': '',
                                'area': ''
                            }
                        )
                        if created:
                            print(f"Created new Person: {person_obj.cedula}")

                        card_data = {
                            'person': person_obj,
                            'tipo_tarjeta': row.get('Tipo de Tarjeta', ''),
                            'numero_tarjeta': str(row.get('Número de Tarjeta', '')),
                            'moneda': row.get('Moneda', ''),
                            'trm_cierre': str(row.get('TRM Cierre', '')),
                            'valor_original': str(row.get('Valor Original', '')),
                            'valor_cop': str(row.get('Valor COP', '')),
                            'numero_autorizacion': str(row.get('Número de Autorización', '')),
                            'fecha_transaccion': pd.to_datetime(row.get('Fecha de Transacción'), errors='coerce').date() if pd.notna(row.get('Fecha de Transacción')) else None,
                            'dia': row.get('Día', ''),
                            'descripcion': row.get('Descripción', ''),
                            'categoria': row.get('Categoría', ''),
                            'subcategoria': row.get('Subcategoría', ''),
                            'zona': row.get('Zona', ''),
                            # No 'cant_tarjetas', 'cargos_abonos', 'archivo', 'cedula_TC' as per new model
                        }

                        # Check for existing transaction to avoid duplicates on re-import
                        lookup_fields = {
                            'person': person_obj,
                            'fecha_transaccion': card_data['fecha_transaccion'],
                            'valor_original': card_data['valor_original']
                        }
                        if card_data['numero_autorizacion'] and card_data['numero_autorizacion'] != 'Sin transacciones':
                            lookup_fields['numero_autorizacion'] = card_data['numero_autorizacion']
                        else:
                            # If no auth number, try to use description to make it somewhat unique
                            lookup_fields['descripcion'] = card_data['descripcion']

                        lookup_fields = {k: v for k, v in lookup_fields.items() if v is not None}


                        try:
                            obj, created = CreditCard.objects.update_or_create( # Changed to CreditCard
                                defaults=card_data,
                                **lookup_fields
                            )
                            if created:
                                transactions_created += 1
                            else:
                                transactions_updated += 1
                        except Exception as e:
                            messages.error(request, f"Error saving transaction row {index} for {cleaned_cedula}: {e}", extra_tags='import_tcs')
                            print(f"Error saving transaction row {index} for {cleaned_cedula}: {e} - Data: {card_data}")

                    messages.success(request, f'Datos de extractos cargados a la base de datos. {transactions_created} creados, {transactions_updated} actualizados.', extra_tags='import_tcs')

                messages.success(request, f'Se procesaron {files_saved} archivos PDF de extractos.', extra_tags='import_tcs')
            except Exception as e:
                messages.error(request, f'Error durante el procesamiento de los PDFs de extractos: {e}', extra_tags='import_tcs')
        else:
            messages.warning(request, 'No se pudieron guardar los archivos PDF para procesar.', extra_tags='import_tcs')

    return redirect('import')


# Function for handling categorias.xlsx upload
def import_categorias(request):
    if request.method == 'POST':
        if 'categorias_excel_file' in request.FILES:
            uploaded_file = request.FILES['categorias_excel_file']
            file_name = "categorias.xlsx"

            target_directory = os.path.join(settings.BASE_DIR, "core", "src")
            os.makedirs(target_directory, exist_ok=True)

            file_path = os.path.join(target_directory, file_name)

            try:
                with open(file_path, 'wb+') as destination:
                    for chunk in uploaded_file.chunks():
                        destination.write(chunk)

                df = pd.read_excel(file_path)
                messages.success(request, f'Archivo "{file_name}" importado correctamente. {len(df)} registros procesados.', extra_tags='import_categorias')
            except Exception as e:
                messages.error(request, f'Error al importar el archivo de categorías: {e}', extra_tags='import_categorias')
        else:
            messages.error(request, 'No se seleccionó ningún archivo de categorías.', extra_tags='import_categorias')
    return redirect('import')
"@

# Create core/conflicts.py
Set-Content -Path "core/conflicts.py" -Value @"
import pandas as pd
import os
from openpyxl.utils import get_column_letter
from openpyxl.styles import numbers

def extract_specific_columns(input_file, output_file, custom_headers=None, year=None):
    try:
        # Setup output directory
        os.makedirs(os.path.dirname(output_file), exist_ok=True)

        # Read raw data (no automatic parsing)
        df = pd.read_excel(input_file, header=None)

        # Check initial number of rows in the raw data (starting from row 4, which is index 3)
        initial_raw_rows = df.shape[0] - 3
        print(f"Initial raw data rows (after header rows): {initial_raw_rows}")

        # Column selection (first 11 + specified extras)
        base_cols = list(range(11))  # Columns 0-10 (A-K)
        # Add the new detail columns
        extra_cols = [11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29]
        selected_cols = [col for col in base_cols + extra_cols if col < df.shape[1]]

        # This operation itself selects all rows from index 3 onwards for the selected columns
        result = df.iloc[3:, selected_cols].copy()
        result.columns = df.iloc[2, selected_cols].values

        print(f"Rows after initial column selection and header application: {result.shape[0]}")

        # Apply custom headers if provided
        if custom_headers is not None:
            if len(custom_headers) != len(result.columns):
                # Revisit this check if 'Año' is added directly into custom_headers list
                raise ValueError(f"Custom headers count ({len(custom_headers)}) doesn't match column count ({len(result.columns)})")
            result.columns = custom_headers
            print(f"Columns after applying custom headers: {result.columns.tolist()}")

        # Add the 'Año' column to the result DataFrame
        if year is not None:
            result['Año'] = year
        else:
            try:
                filename_without_ext = os.path.basename(input_file).split('.')[0]
                year_from_filename = int("".join(filter(str.isdigit, filename_without_ext)))
                result['Año'] = year_from_filename
                print(f"Deduced year from filename: {year_from_filename}")
            except ValueError:
                print("Warning: Could not extract year from filename and no year was provided. 'Año' column will be empty.")
                result['Año'] = pd.NA # Or a default value if preferred

        # Ensure 'Nombre' concatenation handles all rows
        primer_nombre_col_idx = 3 
        primer_apellido_col_idx = 4 
        segundo_apellido_col_idx = 5 

        if primer_nombre_col_idx < df.shape[1] and primer_apellido_col_idx < df.shape[1] and segundo_apellido_col_idx < df.shape[1]:
            temp_df_for_name = df.iloc[3:, [2, 3, 4, 5]].copy()
            temp_df_for_name = temp_df_for_name.fillna('')
            result_nombre_series = temp_df_for_name.iloc[:, 0].astype(str) + " " + \
                                   temp_df_for_name.iloc[:, 1].astype(str) + " " + \
                                   temp_df_for_name.iloc[:, 2].astype(str) + " " + \
                                   temp_df_for_name.iloc[:, 3].astype(str)
            # Ensure the 'Nombre' column is assigned based on the index of 'result'
            result["Nombre"] = result_nombre_series.values # Assign values directly to avoid index alignment issues if result has non-contiguous index
            print(f"Rows after 'Nombre' concatenation: {result.shape[0]}")
        else:
            print("Warning: Not all name columns (C, D, E, F) found in the input DataFrame for name concatenation.")


        # Process "Nombre" column AFTER merging
        if "Nombre" in result.columns:
            result["Nombre"] = result["Nombre"].fillna("")
            result["Nombre"] = result["Nombre"].replace(r'(?i)\bNan\b', '', regex=True)
            result["Nombre"] = result["Nombre"].str.replace(r'\s+', ' ', regex=True).str.strip()
            result["Nombre"] = result["Nombre"].str.title()
            print(f"Rows after 'Nombre' cleanup: {result.shape[0]}")

        # Replace empty strings with pd.NA (NaN)
        result.replace('', pd.NA, inplace=True)
        print(f"Rows after replacing empty strings with NA: {result.shape[0]}")


        # Special handling for Column J (input index 9), which becomes 'Fecha de Inicio' in custom headers
        if "Fecha de Inicio" in result.columns:
            date_col = "Fecha de Inicio"

            result[date_col] = pd.to_datetime(
                result[date_col],
                dayfirst=True,
                errors='coerce'
            )
            print(f"Rows after date conversion: {result.shape[0]}")

            # Save with Excel formatting
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                result.to_excel(writer, index=False)

                worksheet = writer.sheets['Sheet1']
                date_col_letter = get_column_letter(result.columns.get_loc(date_col) + 1)

                for cell in worksheet[date_col_letter]:
                    if cell.row == 1:
                        continue
                    cell.number_format = 'DD/MM/YYYY'

                for idx, col in enumerate(result.columns):
                    col_letter = get_column_letter(idx+1)
                    worksheet.column_dimensions[col_letter].width = max(
                        len(str(col))+2,
                        (result[col].astype(str).str.len().max() or 0) + 2
                    )
            print(f"Successfully saved '{output_file}' with {result.shape[0]} rows.")
        else:
            print("Warning: 'Fecha de Inicio' column not found in processed data. Saving without date formatting.")
            result.to_excel(output_file, index=False)
            print(f"Successfully saved '{output_file}' with {result.shape[0]} rows (no date formatting).")

        return result

    except Exception as e:
        print(f"Error in extract_specific_columns: {str(e)}")
        return pd.DataFrame()

def generate_justrue_file(input_df, output_file):
    try:
        os.makedirs(os.path.dirname(output_file), exist_ok=True)

        justrue_data = pd.DataFrame()

        if 'Cedula' in input_df.columns:
            justrue_data['Cedula'] = input_df['Cedula']
        else:
            print("Error: 'Cedula' column not found in the input DataFrame for jusTrue file generation.")
            return

        if 'Año' in input_df.columns:
            justrue_data['Año'] = input_df['Año']

        q_columns = [f"Q{i}" for i in range(1, 8)] + [f"Q{i}" for i in range(10, 12)]

        for q_col in q_columns:
            if q_col in input_df.columns:
                # Assign 'true' where the condition is met, keeping original index alignment
                # Use .loc for setting values to avoid SettingWithCopyWarning
                justrue_data[q_col] = input_df[q_col].apply(lambda x: 'true' if str(x).lower() == 'true' else pd.NA)
            else:
                print(f"Warning: Column '{q_col}' not found in the input DataFrame for jusTrue file generation.")

        cols_to_check = [f"Q{i}" for i in range(1, 8)]

        initial_justrue_rows = justrue_data.shape[0]
        justrue_data.dropna(subset=cols_to_check, how='all', inplace=True)
        print(f"Rows in jusTrue data before dropping NAs: {initial_justrue_rows}")
        print(f"Rows in jusTrue data after dropping NAs (only rows with at least one Q1-Q7 'true'): {justrue_data.shape[0]}")


        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            justrue_data.to_excel(writer, index=False)

            worksheet = writer.sheets['Sheet1']
            for idx, col in enumerate(justrue_data.columns):
                col_letter = get_column_letter(idx + 1)
                worksheet.column_dimensions[col_letter].width = max(
                    len(str(col)) + 2,
                    (justrue_data[col].astype(str).str.len().max() or 0) + 2
                )

        print(f"Successfully created '{output_file}' with filtered data.")

    except Exception as e:
        print(f"Error creating jusTrue file: {str(e)}")

# 'Año' will be added dynamically, so it's not in this list.
custom_headers = [
    "ID", "Cedula", "Nombre", "1er Nombre", "1er Apellido",
    "2do Apellido", "Compañía", "Cargo", "Email", "Fecha de Inicio",
    "Q1", "Q1 Detalle", "Q2", "Q2 Detalle", "Q3", "Q3 Detalle",
    "Q4", "Q4 Detalle", "Q5", "Q5 Detalle", "Q6", "Q6 Detalle",
    "Q7", "Q7 Detalle", "Q8", "Q9", "Q10", "Q10 Detalle", "Q11", "Q11 Detalle"
]

# Assuming current year is 2024 for the conflictos.xlsx file.
current_year = 2024

processed_df = extract_specific_columns(
    input_file="core/src/conflictos.xlsx",
    output_file="core/src/conflicts.xlsx",
    custom_headers=custom_headers,
    year=current_year # Pass the year explicitly
)

# Then, if the processing was successful, generate the jusTrue.xlsx file
if not processed_df.empty:
    generate_justrue_file(
        input_df=processed_df,
        output_file="core/src/jusTrue.xlsx"
    )
"@

# Create cats.py
Set-Content -Path "core/cats.py" -Value @"
import pandas as pd
from datetime import datetime

# Shared constants and functions
TRM_DICT = {
    2020: 3432.50,
    2021: 3981.16,
    2022: 4810.20,
    2023: 4780.38,
    2024: 4409.00
}

CURRENCY_RATES = {
    2020: {
        'EUR': 1.141, 'GBP': 1.280, 'AUD': 0.690, 'CAD': 0.746,
        'HNL': 0.0406, 'AWG': 0.558, 'DOP': 0.0172, 'PAB': 1.000,
        'CLP': 0.00126, 'CRC': 0.00163, 'ARS': 0.0119, 'ANG': 0.558,
        'COP': 0.00026,  'BBD': 0.50, 'MXN': 0.0477, 'BOB': 0.144, 'BSD': 1.00,
        'GYD': 0.0048, 'UYU': 0.025, 'DKK': 0.146, 'KYD': 1.20, 'BMD': 1.00, 
        'VEB': 0.0000000248, 'VES': 0.000000248, 'BRL': 0.187, 'NIO': 0.0278
    },
    2021: {
        'EUR': 1.183, 'GBP': 1.376, 'AUD': 0.727, 'CAD': 0.797,
        'HNL': 0.0415, 'AWG': 0.558, 'DOP': 0.0176, 'PAB': 1.000,
        'CLP': 0.00118, 'CRC': 0.00156, 'ARS': 0.00973, 'ANG': 0.558,
        'COP': 0.00027, 'BBD': 0.50, 'MXN': 0.0492, 'BOB': 0.141, 'BSD': 1.00,
        'GYD': 0.0047, 'UYU': 0.024, 'DKK': 0.155, 'KYD': 1.20, 'BMD': 1.00,
        'VEB': 0.00000000002, 'VES': 0.00000002, 'BRL': 0.192, 'NIO': 0.0285
    },
    2022: {
        'EUR': 1.051, 'GBP': 1.209, 'AUD': 0.688, 'CAD': 0.764,
        'HNL': 0.0408, 'AWG': 0.558, 'DOP': 0.0181, 'PAB': 1.000,
        'CLP': 0.00117, 'CRC': 0.00155, 'ARS': 0.00597, 'ANG': 0.558,
        'COP': 0.00021, 'BBD': 0.50, 'MXN': 0.0497, 'BOB': 0.141, 'BSD': 1.00,
        'GYD': 0.0047, 'UYU': 0.025, 'DKK': 0.141, 'KYD': 1.20, 'BMD': 1.00,
        'VEB': 0, 'VES': 0.000000001, 'BRL': 0.196, 'NIO': 0.0267
    },
    2023: {
        'EUR': 1.096, 'GBP': 1.264, 'AUD': 0.676, 'CAD': 0.741,
        'HNL': 0.0406, 'AWG': 0.558, 'DOP': 0.0177, 'PAB': 1.000,
        'CLP': 0.00121, 'CRC': 0.00187, 'ARS': 0.00275, 'ANG': 0.558,
        'COP': 0.00022, 'BBD': 0.50, 'MXN': 0.0564, 'BOB': 0.143, 'BSD': 1.00,
        'GYD': 0.0047, 'UYU': 0.025, 'DKK': 0.148, 'KYD': 1.20, 'BMD': 1.00,
        'VEB': 0, 'VES': 0.000000001, 'BRL': 0.194, 'NIO': 0.0267
    },
    2024: {
        'EUR': 1.093, 'GBP': 1.267, 'AUD': 0.674, 'CAD': 0.742,
        'HNL': 0.0405, 'AWG': 0.558, 'DOP': 0.0170, 'PAB': 1.000,
        'CLP': 0.00111, 'CRC': 0.00192, 'ARS': 0.00121, 'ANG': 0.558,
        'COP': 0.00022, 'BBD': 0.50, 'MXN': 0.0547, 'BOB': 0.142, 'BSD': 1.00,
        'GYD': 0.0047, 'UYU': 0.024, 'DKK': 0.147, 'KYD': 1.20, 'BMD': 1.00,
        'VEB': 0, 'VES': 0.000000001, 'BRL': 0.190, 'NIO': 0.0260 }
}

# Define the periodo dataframe internally
PERIODO_DF = pd.DataFrame({
    'Id': [2, 6, 7, 8],
    'Activo': [True, True, True, True],
    'Año': ['Friday, January 01, 2021', 'Saturday, January 01, 2022', 
            'Sunday, January 01, 2023', 'Monday, January 01, 2024'],
    'FechaFinDeclaracion': ['4/30/2022', '3/31/2023', '5/12/2024', '1/1/2025'],
    'FechaInicioDeclaracion': ['6/1/2021', '10/19/2022', '11/1/2023', '10/2/2024'],
    'Año declaracion': ['2,021', '2,022', '2,023', '2,024']
})

def get_trm(year):
    """Gets TRM for a given year from the dictionary"""
    return TRM_DICT.get(year)

def get_exchange_rate(currency_code, year):
    """Gets exchange rate for a given currency and year from the dictionary"""
    year_rates = CURRENCY_RATES.get(year)
    if year_rates:
        return year_rates.get(currency_code)
    return None

def get_currency_code(moneda_text):
    """Extracts the currency code from the 'Texto Moneda' field"""
    currency_mapping = {
        'HNL -Lempira hondureño': 'HNL',
        'EUR - Euro': 'EUR',
        'AWG - Florín holandés o de Aruba': 'AWG',
        'DOP - Peso dominicano': 'DOP',
        'PAB -Balboa panameña': 'PAB', 
        'CLP - Peso chileno': 'CLP',
        'CRC - Colón costarricense': 'CRC',
        'ARS - Peso argentino': 'ARS',
        'AUD - Dólar australiano': 'AUD',
        'ANG - Florín holandés': 'ANG',
        'CAD -Dólar canadiense': 'CAD',
        'GBP - Libra esterlina': 'GBP',
        'USD - Dolar estadounidense': 'USD',
        'COP - Peso colombiano': 'COP',
        'BBD - Dólar de Barbados o Baja': 'BBD',
        'MXN - Peso mexicano': 'MXN',
        'BOB - Boliviano': 'BOB',
        'BSD - Dolar bahameño': 'BSD',
        'GYD - Dólar guyanés': 'GYD',
        'UYU - Peso uruguayo': 'UYU',
        'DKK - Corona danesa': 'DKK',
        'KYD - Dólar de las Caimanes': 'KYD',
        'BMD - Dólar de las Bermudas': 'BMD',
        'VEB - Bolívar venezolano': 'VEB',  
        'VES - Bolívar soberano': 'VES',  
        'BRL - Real brasilero': 'BRL',  
        'NIO - Córdoba nicaragüense': 'NIO',
    }
    return currency_mapping.get(moneda_text)

def get_valid_year(row, periodo_df=None):
    """Extracts a valid year, handling missing values and format variations."""
    try:
        fkIdPeriodo = pd.to_numeric(row['fkIdPeriodo'], errors='coerce')
        if pd.isna(fkIdPeriodo):  # Handle missing fkIdPeriodo
            print(f"Warning: Missing fkIdPeriodo at index {row.name}. Skipping row.")
            return None

        # Use the internal PERIODO_DF if no periodo_df is provided
        periodo_df = periodo_df if periodo_df is not None else PERIODO_DF
        
        matching_row = periodo_df[periodo_df['Id'] == fkIdPeriodo]
        if matching_row.empty:
            print(f"Warning: No matching Id found in periodo data for fkIdPeriodo {fkIdPeriodo} at index {row.name}. Skipping row.")
            return None

        year_str = matching_row['Año declaracion'].iloc[0]

        try:
            # Clean the year string by removing commas and converting to integer
            year = int(year_str.replace(',', ''))
            return year
        except (ValueError, TypeError):
            try:
                year = pd.to_datetime(year_str, errors='coerce').year
                if pd.isna(year):
                    raise ValueError
                return year
            except ValueError:
                print(f"Warning: Invalid year format '{year_str}' for fkIdPeriodo {fkIdPeriodo} at index {row.name}. Skipping row.")
                return None

    except Exception as e:
        print(f"Error in get_valid_year for fkIdPeriodo {fkIdPeriodo} at index {row.name}: {e}")
        return None

def analyze_banks(file_path, output_file_path, periodo_file_path=None):
    """Analyze bank account data"""
    df = pd.read_excel(file_path)
    
    maintain_columns = [
        'fkIdPeriodo', 'fkIdEstado',
        'Año Creación', 'Año Envío', 'Usuario',
        'Nombre', 'Compañía', 'Cargo', 'RUBRO DE DECLARACIÓN', 'fkIdDeclaracion',
        'Banco - Entidad', 'Banco - Tipo Cuenta', 'Texto Moneda',
        'Banco - fkIdPaís', 'Banco - Nombre País',
        'Banco - Saldo', 'Banco - Comentario'
    ]
    
    banks_df = df.loc[df['RUBRO DE DECLARACIÓN'] == 'Banco', maintain_columns].copy()
    banks_df = banks_df[banks_df['fkIdEstado'] != 1]
    
    banks_df['Banco - Saldo COP'] = 0.0
    banks_df['TRM Aplicada'] = None
    banks_df['Tasa USD'] = None
    banks_df['Año Declaración'] = None 
    
    for index, row in banks_df.iterrows():
        try:
            year = get_valid_year(row)
            if year is None:
                print(f"Warning: Could not determine valid year for index {index} and fkIdPeriodo {row['fkIdPeriodo']}. Skipping row.")
                banks_df.loc[index, 'Año Declaración'] = "Año no encontrado"
                continue 
                
            banks_df.loc[index, 'Año Declaración'] = year
            currency_code = get_currency_code(row['Texto Moneda'])
            
            if currency_code == 'COP':
                banks_df.loc[index, 'Banco - Saldo COP'] = float(row['Banco - Saldo'])
                banks_df.loc[index, 'TRM Aplicada'] = 1.0
                banks_df.loc[index, 'Tasa USD'] = None
                continue
                
            if currency_code:
                trm = get_trm(year)
                usd_rate = 1.0 if currency_code == 'USD' else get_exchange_rate(currency_code, year)
                
                if trm and usd_rate:
                    usd_amount = float(row['Banco - Saldo']) * usd_rate
                    cop_amount = usd_amount * trm
                    
                    banks_df.loc[index, 'Banco - Saldo COP'] = cop_amount
                    banks_df.loc[index, 'TRM Aplicada'] = trm
                    banks_df.loc[index, 'Tasa USD'] = usd_rate
                else:
                    print(f"Warning: Missing conversion rates for {currency_code} in {year} at index {index}")
            else:
                print(f"Warning: Unknown currency format '{row['Texto Moneda']}' at index {index}")
                
        except Exception as e:
            print(f"Warning: Error processing row at index {index}: {e}")
            banks_df.loc[index, 'Año Declaración'] = "Error de procesamiento"
            continue
    
    banks_df.to_excel(output_file_path, index=False)

def analyze_debts(file_path, output_file_path, periodo_file_path=None):
    """Analyze debts data"""
    df = pd.read_excel(file_path)
    
    maintain_columns = [
        'fkIdPeriodo', 'fkIdEstado',
        'Año Creación', 'Año Envío', 'Usuario', 'Nombre',
        'Compañía', 'Cargo', 'RUBRO DE DECLARACIÓN', 'fkIdDeclaracion',
        'Pasivos - Entidad Personas',
        'Pasivos - Tipo Obligación', 'fkIdMoneda', 'Texto Moneda',
        'Pasivos - Valor', 'Pasivos - Comentario', 'Pasivos - Valor COP'
    ]
    
    debts_df = df.loc[df['RUBRO DE DECLARACIÓN'] == 'Pasivo', maintain_columns].copy()
    debts_df = debts_df[debts_df['fkIdEstado'] != 1]
    
    debts_df['Pasivos - Valor COP'] = 0.0
    debts_df['TRM Aplicada'] = None
    debts_df['Tasa USD'] = None
    debts_df['Año Declaración'] = None 
    
    for index, row in debts_df.iterrows():
        try:
            year = get_valid_year(row)
            if year is None:
                print(f"Warning: Could not determine valid year for index {index}. Skipping row.")
                continue
            
            debts_df.loc[index, 'Año Declaración'] = year
            currency_code = get_currency_code(row['Texto Moneda'])
            
            if currency_code == 'COP':
                debts_df.loc[index, 'Pasivos - Valor COP'] = float(row['Pasivos - Valor'])
                debts_df.loc[index, 'TRM Aplicada'] = 1.0
                debts_df.loc[index, 'Tasa USD'] = None
                continue
            
            if currency_code:
                trm = get_trm(year)
                usd_rate = 1.0 if currency_code == 'USD' else get_exchange_rate(currency_code, year)
                
                if trm and usd_rate:
                    usd_amount = float(row['Pasivos - Valor']) * usd_rate
                    cop_amount = usd_amount * trm
                    
                    debts_df.loc[index, 'Pasivos - Valor COP'] = cop_amount
                    debts_df.loc[index, 'TRM Aplicada'] = trm
                    debts_df.loc[index, 'Tasa USD'] = usd_rate
                else:
                    print(f"Warning: Missing conversion rates for {currency_code} in {year} at index {index}")
            else:
                print(f"Warning: Unknown currency format '{row['Texto Moneda']}' at index {index}")
                
        except Exception as e:
            print(f"Warning: Error processing row at index {index}: {e}")
            debts_df.loc[index, 'Año Declaración'] = "Error de procesamiento"
            continue

    debts_df.to_excel(output_file_path, index=False)

def analyze_goods(file_path, output_file_path, periodo_file_path=None):
    """Analyze goods/patrimony data"""
    df = pd.read_excel(file_path)
    
    maintain_columns = [
        'fkIdPeriodo', 'fkIdEstado',
        'Año Creación', 'Año Envío', 'Usuario', 'Nombre',
        'Compañía', 'Cargo', 'RUBRO DE DECLARACIÓN', 'fkIdDeclaracion',
        'Patrimonio - Activo', 'Patrimonio - % Propiedad',
        'Patrimonio - Propietario', 'Patrimonio - Valor Comercial',
        'Patrimonio - Comentario',
        'Patrimonio - Valor Comercial COP', 'Texto Moneda'
    ]
    
    goods_df = df.loc[df['RUBRO DE DECLARACIÓN'] == 'Patrimonio', maintain_columns].copy()
    goods_df = goods_df[goods_df['fkIdEstado'] != 1]
    
    goods_df['Patrimonio - Valor COP'] = 0.0
    goods_df['TRM Aplicada'] = None
    goods_df['Tasa USD'] = None
    goods_df['Año Declaración'] = None 
    
    for index, row in goods_df.iterrows():
        try:
            year = get_valid_year(row)
            if year is None:
                print(f"Warning: Could not determine valid year for index {index}. Skipping row.")
                continue
                
            goods_df.loc[index, 'Año Declaración'] = year
            currency_code = get_currency_code(row['Texto Moneda'])
            
            if currency_code == 'COP':
                goods_df.loc[index, 'Patrimonio - Valor COP'] = float(row['Patrimonio - Valor Comercial'])
                goods_df.loc[index, 'TRM Aplicada'] = 1.0
                goods_df.loc[index, 'Tasa USD'] = None
                continue
            
            if currency_code:
                trm = get_trm(year)
                usd_rate = 1.0 if currency_code == 'USD' else get_exchange_rate(currency_code, year)
                
                if trm and usd_rate:
                    usd_amount = float(row['Patrimonio - Valor Comercial']) * usd_rate
                    cop_amount = usd_amount * trm
                    
                    goods_df.loc[index, 'Patrimonio - Valor COP'] = cop_amount
                    goods_df.loc[index, 'TRM Aplicada'] = trm
                    goods_df.loc[index, 'Tasa USD'] = usd_rate
                else:
                    print(f"Warning: Missing conversion rates for {currency_code} in {year} at index {index}")
            else:
                print(f"Warning: Unknown currency format '{row['Texto Moneda']}' at index {index}")
                
        except Exception as e:
            print(f"Warning: Error processing row at index {index}: {e}")
            continue
        
    goods_df['Patrimonio - Valor Corregido'] = goods_df['Patrimonio - Valor COP'] * (goods_df['Patrimonio - % Propiedad'] / 100)
    
    # Rename columns for consistency
    rename_dict = {
        'Patrimonio - Valor Corregido': 'Bienes - Valor Corregido',
        'Patrimonio - Valor Comercial COP': 'Bienes - Valor Comercial COP',
        'Patrimonio - Comentario': 'Bienes - Comentario',
        'Patrimonio - Valor Comercial': 'Bienes - Valor Comercial',
        'Patrimonio - Propietario': 'Bienes - Propietario',
        'Patrimonio - % Propiedad': 'Bienes - % Propiedad',
        'Patrimonio - Activo': 'Bienes - Activo',
        'Patrimonio - Valor COP': 'Bienes - Valor COP'
    }
    goods_df = goods_df.rename(columns=rename_dict)
    
    goods_df.to_excel(output_file_path, index=False)

def analyze_incomes(file_path, output_file_path, periodo_file_path=None):
    """Analyze income data"""
    df = pd.read_excel(file_path)
    
    maintain_columns = [
        'fkIdPeriodo', 'fkIdEstado',
        'Año Creación', 'Año Envío', 'Usuario', 'Nombre',
        'Compañía', 'Cargo', 'RUBRO DE DECLARACIÓN', 'fkIdDeclaracion',
        'Ingresos - fkIdConcepto', 'Ingresos - Texto Concepto',
        'Ingresos - Valor', 'Ingresos - Comentario', 'Ingresos - Otros',
        'Ingresos - Valor_COP', 'Texto Moneda'
    ]

    incomes_df = df.loc[df['RUBRO DE DECLARACIÓN'] == 'Ingreso', maintain_columns].copy()
    incomes_df = incomes_df[incomes_df['fkIdEstado'] != 1]
    
    incomes_df['Ingresos - Valor COP'] = 0.0
    incomes_df['TRM Aplicada'] = None
    incomes_df['Tasa USD'] = None
    incomes_df['Año Declaración'] = None 
    
    for index, row in incomes_df.iterrows():
        try:
            year = get_valid_year(row)
            if year is None:
                print(f"Warning: Could not determine valid year for index {index}. Skipping row.")
                continue
            
            incomes_df.loc[index, 'Año Declaración'] = year
            currency_code = get_currency_code(row['Texto Moneda'])
            
            if currency_code == 'COP':
                incomes_df.loc[index, 'Ingresos - Valor COP'] = float(row['Ingresos - Valor'])
                incomes_df.loc[index, 'TRM Aplicada'] = 1.0
                incomes_df.loc[index, 'Tasa USD'] = None
                continue
            
            if currency_code:
                trm = get_trm(year)
                usd_rate = 1.0 if currency_code == 'USD' else get_exchange_rate(currency_code, year)
                
                if trm and usd_rate:
                    usd_amount = float(row['Ingresos - Valor']) * usd_rate
                    cop_amount = usd_amount * trm
                    
                    incomes_df.loc[index, 'Ingresos - Valor COP'] = cop_amount
                    incomes_df.loc[index, 'TRM Aplicada'] = trm
                    incomes_df.loc[index, 'Tasa USD'] = usd_rate
                else:
                    print(f"Warning: Missing conversion rates for {currency_code} in {year} at index {index}")
            else:
                print(f"Warning: Unknown currency format '{row['Texto Moneda']}' at index {index}")
                
        except Exception as e:
            print(f"Warning: Error processing row at index {index}: {e}")
            continue
    
    incomes_df.to_excel(output_file_path, index=False)

def analyze_investments(file_path, output_file_path, periodo_file_path=None):
    """Analyze investment data"""
    df = pd.read_excel(file_path)
    
    maintain_columns = [
        'fkIdPeriodo', 'fkIdEstado',
        'Año Creación', 'Año Envío', 'Usuario', 'Nombre',
        'Compañía', 'Cargo', 'RUBRO DE DECLARACIÓN', 'fkIdDeclaracion',
        'Inversiones - Tipo Inversión', 'Inversiones - Entidad',
        'Inversiones - Valor', 'Inversiones - Comentario',
        'Inversiones - Valor COP', 'Texto Moneda'
    ]
    
    invest_df = df.loc[df['RUBRO DE DECLARACIÓN'] == 'Inversión', maintain_columns].copy()
    invest_df = invest_df[invest_df['fkIdEstado'] != 1]
    
    invest_df['Inversiones - Valor COP'] = 0.0
    invest_df['TRM Aplicada'] = None
    invest_df['Tasa USD'] = None
    invest_df['Año Declaración'] = None 
    
    for index, row in invest_df.iterrows():
        try:
            year = get_valid_year(row)
            if year is None:
                print(f"Warning: Could not determine valid year for index {index}. Skipping row.")
                continue
            
            invest_df.loc[index, 'Año Declaración'] = year
            currency_code = get_currency_code(row['Texto Moneda'])
            
            if currency_code == 'COP':
                invest_df.loc[index, 'Inversiones - Valor COP'] = float(row['Inversiones - Valor'])
                invest_df.loc[index, 'TRM Aplicada'] = 1.0
                invest_df.loc[index, 'Tasa USD'] = None
                continue
            
            if currency_code:
                trm = get_trm(year)
                usd_rate = 1.0 if currency_code == 'USD' else get_exchange_rate(currency_code, year)
                
                if trm and usd_rate:
                    usd_amount = float(row['Inversiones - Valor']) * usd_rate
                    cop_amount = usd_amount * trm
                    
                    invest_df.loc[index, 'Inversiones - Valor COP'] = cop_amount
                    invest_df.loc[index, 'TRM Aplicada'] = trm
                    invest_df.loc[index, 'Tasa USD'] = usd_rate
                else:
                    print(f"Warning: Missing conversion rates for {currency_code} in {year} at index {index}")
            else:
                print(f"Warning: Unknown currency format '{row['Texto Moneda']}' at index {index}")
                
        except Exception as e:
            print(f"Warning: Error processing row at index {index}: {e}")
            continue
    
    invest_df.to_excel(output_file_path, index=False)

def run_all_analyses():
    """Run all analysis functions with their respective file paths"""
    file_path = 'core/src/data.xlsx'
    
    analyze_banks(file_path, 'core/src/banks.xlsx')
    analyze_debts(file_path, 'core/src/debts.xlsx')
    analyze_goods(file_path, 'core/src/goods.xlsx')
    analyze_incomes(file_path, 'core/src/incomes.xlsx')
    analyze_investments(file_path, 'core/src/investments.xlsx')

if __name__ == "__main__":
    run_all_analyses()
"@

# Create nets.py
Set-Content -Path "core/nets.py" -Value @"
import pandas as pd

# Common columns used across all analyses
COMMON_COLUMNS = [
    'Usuario', 'Nombre', 'Compañía', 'Cargo',
    'fkIdPeriodo', 'fkIdEstado',
    'Año Creación', 'Año Envío',
    'RUBRO DE DECLARACIÓN', 'fkIdDeclaracion',
    'Año Declaración'
]

# Base groupby columns for summaries
BASE_GROUPBY = ['Usuario', 'Nombre', 'Compañía', 'Cargo', 'fkIdPeriodo', 'Año Declaración', 'Año Creación']

def analyze_banks(file_path, output_file_path):
    """Analyze bank accounts data"""
    df = pd.read_excel(file_path)

    # Specific columns for banks
    bank_columns = [
        'Banco - Entidad', 'Banco - Tipo Cuenta',
        'Banco - fkIdPaís', 'Banco - Nombre País',
        'Banco - Saldo', 'Banco - Comentario',
        'Banco - Saldo COP'
    ]
    
    df = df[COMMON_COLUMNS + bank_columns]
    
    # Create a temporary combination column for counting
    df_temp = df.copy()
    df_temp['Bank_Account_Combo'] = df['Banco - Entidad'] + "|" + df['Banco - Tipo Cuenta']
    
    # Perform all aggregations
    summary = df_temp.groupby(BASE_GROUPBY).agg(
        **{
            'Cant_Bancos': pd.NamedAgg(column='Banco - Entidad', aggfunc='nunique'),
            'Cant_Cuentas': pd.NamedAgg(column='Bank_Account_Combo', aggfunc='nunique'),
            'Banco - Saldo COP': pd.NamedAgg(column='Banco - Saldo COP', aggfunc='sum')
        }
    ).reset_index()

    summary.to_excel(output_file_path, index=False)
    return summary

def analyze_debts(file_path, output_file_path):
    """Analyze debts data"""
    df = pd.read_excel(file_path)

    # Specific columns for debts
    debt_columns = [
        'Pasivos - Entidad Personas', 'Pasivos - Tipo Obligación', 
        'Pasivos - Valor', 'Pasivos - Comentario',
        'Pasivos - Valor COP', 'Texto Moneda'
    ]
    
    df = df[COMMON_COLUMNS + debt_columns]
    
    # Calculate total Pasivos and count occurrences
    summary = df.groupby(BASE_GROUPBY).agg({      
        'Pasivos - Valor COP': 'sum',
        'Pasivos - Entidad Personas': 'count'
    }).reset_index()

    # Rename columns for clarity
    summary = summary.rename(columns={
        'Pasivos - Entidad Personas': 'Cant_Deudas',
        'Pasivos - Valor COP': 'Total Pasivos'
    })

    summary.to_excel(output_file_path, index=False)
    return summary

def analyze_goods(file_path, output_file_path):
    """Analyze goods/assets data"""
    df = pd.read_excel(file_path)
    
    # Specific columns for goods
    goods_columns = [
        'Bienes - Activo', 'Bienes - % Propiedad',
        'Bienes - Propietario', 'Bienes - Valor Comercial',
        'Bienes - Comentario', 'Bienes - Valor Comercial COP',
        'Bienes - Valor Corregido'
    ]
    
    df = df[COMMON_COLUMNS + goods_columns]

    summary = df.groupby(BASE_GROUPBY).agg({
        'Bienes - Valor Corregido': 'sum',
        'Bienes - Activo': 'count' 
    }).reset_index()

    # Rename columns for clarity
    summary = summary.rename(columns={
        'Bienes - Activo': 'Cant_Bienes',
        'Bienes - Valor Corregido': 'Total Bienes'
    })

    summary.to_excel(output_file_path, index=False) 
    return summary

def analyze_incomes(file_path, output_file_path):
    """Analyze income data"""
    df = pd.read_excel(file_path)
    
    # Specific columns for incomes
    income_columns = [
        'Ingresos - fkIdConcepto', 'Ingresos - Texto Concepto',
        'Ingresos - Valor', 'Ingresos - Comentario',
        'Ingresos - Otros', 'Ingresos - Valor COP',
        'Texto Moneda'
    ]

    df = df[COMMON_COLUMNS + income_columns]
    
    # Calculate Ingresos and count occurrences
    summary = df.groupby(BASE_GROUPBY).agg({
        'Ingresos - Valor COP': 'sum',
        'Ingresos - Texto Concepto': 'count'
    }).reset_index()

    # Rename columns for clarity
    summary = summary.rename(columns={
        'Ingresos - Texto Concepto': 'Cant_Ingresos',
        'Ingresos - Valor COP': 'Total Ingresos'
    })

    summary.to_excel(output_file_path, index=False)
    return summary

def analyze_investments(file_path, output_file_path):
    """Analyze investments data"""
    df = pd.read_excel(file_path)
    
    # Specific columns for investments
    invest_columns = [
        'Inversiones - Tipo Inversión', 'Inversiones - Entidad',
        'Inversiones - Valor', 'Inversiones - Comentario',
        'Inversiones - Valor COP', 'Texto Moneda'
    ]
    
    df = df[COMMON_COLUMNS + invest_columns]
    
    # Calculate total Inversiones and count occurrences
    summary = df.groupby(BASE_GROUPBY + ['Inversiones - Tipo Inversión']).agg( 
        {'Inversiones - Valor COP': 'sum',
         'Inversiones - Tipo Inversión': 'count'}
    ).rename(columns={
        'Inversiones - Tipo Inversión': 'Cant_Inversiones',
        'Inversiones - Valor COP': 'Total Inversiones'
    }).reset_index()
    
    summary.to_excel(output_file_path, index=False)
    return summary 

def calculate_assets(banks_file, goods_file, invests_file, output_file):
    """Calculate total assets by combining banks, goods and investments"""
    banks = pd.read_excel(banks_file)
    goods = pd.read_excel(goods_file)
    invests = pd.read_excel(invests_file)

    # Group investments by base columns (summing across types)
    invests_grouped = invests.groupby(BASE_GROUPBY).agg({
        'Total Inversiones': 'sum',
        'Cant_Inversiones': 'sum'
    }).reset_index()

    # Merge all three dataframes
    merged = pd.merge(goods, banks, on=BASE_GROUPBY, how='outer')
    merged = pd.merge(merged, invests_grouped, on=BASE_GROUPBY, how='outer')
    merged.fillna(0, inplace=True)

    # Calculate total assets
    merged['Total Activos'] = (
        merged['Total Bienes'] + 
        merged['Banco - Saldo COP'] + 
        merged['Total Inversiones']
    )

    # Reorder and rename columns
    final_columns = BASE_GROUPBY + [
        'Total Bienes', 'Cant_Bienes',
        'Banco - Saldo COP', 'Cant_Bancos', 'Cant_Cuentas',
        'Total Inversiones', 'Cant_Inversiones',
        'Total Activos'
    ]
    merged = merged[final_columns]

    merged.to_excel(output_file, index=False)
    return merged

def calculate_net_worth(debts_file, assets_file, output_file):
    """Calculate net worth by combining assets and debts"""
    debts = pd.read_excel(debts_file)
    assets = pd.read_excel(assets_file)

    # Merge the summaries
    merged = pd.merge(
        assets, 
        debts, 
        on=BASE_GROUPBY, 
        how='outer'
    )
    merged.fillna(0, inplace=True)
    
    # Calculate net worth
    merged['Total Patrimonio'] = merged['Total Activos'] - merged['Total Pasivos']
    
    # Final column order
    final_columns = BASE_GROUPBY + [
        'Total Activos',
        'Cant_Bienes',
        'Cant_Bancos',
        'Cant_Cuentas',
        'Cant_Inversiones',
        'Total Pasivos',
        'Cant_Deudas',
        'Total Patrimonio'
    ]
    merged = merged[final_columns]
    
    merged.to_excel(output_file, index=False)
    return merged

def run_all_analyses():
    """Run all analyses in sequence with default file paths"""
    # Individual analyses
    bank_summary = analyze_banks(
        'core/src/banks.xlsx',
        'core/src/bankNets.xlsx'
    )
    
    debt_summary = analyze_debts(
        'core/src/debts.xlsx',
        'core/src/debtNets.xlsx'
    )
    
    goods_summary = analyze_goods(
        'core/src/goods.xlsx',
        'core/src/goodNets.xlsx'
    )
    
    income_summary = analyze_incomes(
        'core/src/incomes.xlsx',
        'core/src/incomeNets.xlsx'
    )
    
    invest_summary = analyze_investments(
        'core/src/investments.xlsx',
        'core/src/investNets.xlsx'
    )
    
    # Combined analyses
    assets_summary = calculate_assets(
        'core/src/bankNets.xlsx',
        'core/src/goodNets.xlsx',
        'core/src/investNets.xlsx',
        'core/src/assetNets.xlsx'
    )
    
    net_worth_summary = calculate_net_worth(
        'core/src/debtNets.xlsx',
        'core/src/assetNets.xlsx',
        'core/src/worthNets.xlsx'
    )
    
    return {
        'bank_summary': bank_summary,
        'debt_summary': debt_summary,
        'goods_summary': goods_summary,
        'income_summary': income_summary,
        'invest_summary': invest_summary,
        'assets_summary': assets_summary,
        'net_worth_summary': net_worth_summary
    }

if __name__ == '__main__':
    # Run all analyses when script is executed
    results = run_all_analyses()
    print("All nets analyses completed successfully!")
"@

# Create trends.py
Set-Content -Path "core/trends.py" -Value @"
import pandas as pd

def get_trend_symbol(value):
    """Determine the trend symbol based on the percentage change."""
    try:
        # Check if the value is "N/A" or empty, indicating no trend for the first year
        if value in ["N/A", "0.00%", ""]:
            return "" # Return empty string for no trend
            
        value_float = float(value.strip('%')) / 100
        if pd.isna(value_float):
            return "➡️"
        elif value_float > 0.1:  # more than 10% increase
            return "📈"
        elif value_float < -0.1:  # more than 10% decrease
            return "📉"
        else:
            return "➡️"  # relatively stable
    except Exception:
        return "➡️"

def calculate_variation(df, column):
    """Calculate absolute and relative variations for a specific column."""
    df = df.sort_values(by=['Usuario', 'Año Declaración'])
    
    absolute_col = f'{column} Var. Abs.'
    relative_col = f'{column} Var. Rel.'
    
    df[absolute_col] = df.groupby('Usuario')[column].diff()
    
    # Calculate percentage change
    pct_change = df.groupby('Usuario')[column].pct_change(fill_method=None) * 100
    
    # Apply formatting: "0.00%" for non-NaN values, empty string for NaN (first year)
    df[relative_col] = pct_change.apply(lambda x: f"{x:.2f}%" if pd.notna(x) else "")
    
    return df

def embed_trend_symbols(df, columns):
    """Add trend symbols to variation columns."""
    for col in columns:
        absolute_col = f'{col} Var. Abs.'
        relative_col = f'{col} Var. Rel.'
        
        if absolute_col in df.columns:
            df[absolute_col] = df.apply(
                lambda row: f"{row[absolute_col]:.2f} {get_trend_symbol(row[relative_col])}" 
                if pd.notna(row[f'{col} Var. Abs. No_Symbol']) else "N/A", # Use a temporary column to check for original NaN
                axis=1
            )
        
        if relative_col in df.columns:
            # Check if the underlying relative change was NaN before formatting
            df[relative_col] = df.apply(
                lambda row: f"{row[relative_col]} {get_trend_symbol(row[relative_col])}" if row[relative_col] != "" else "",
                axis=1
            )
    
    return df

def calculate_leverage(df):
    """Calculate financial leverage."""
    df['Apalancamiento'] = (df['Patrimonio'] / df['Activos']) * 100
    return df

def calculate_debt_level(df):
    """Calculate debt level."""
    df['Endeudamiento'] = (df['Pasivos'] / df['Activos']) * 100
    return df

def process_asset_data(df_assets):
    """Process asset data with variations and trends."""
    df_assets_grouped = df_assets.groupby(['Usuario', 'Año Declaración']).agg(
        Banco_Saldo=('Banco - Saldo COP', 'sum'),
        Bienes=('Total Bienes', 'sum'),
        Inversiones=('Total Inversiones', 'sum')
    ).reset_index()

    for column in ['Banco_Saldo', 'Bienes', 'Inversiones']:
        df_assets_grouped = calculate_variation(df_assets_grouped, column)
        # Create a temporary column to hold the absolute variation without symbol for the N/A check
        df_assets_grouped[f'{column} Var. Abs. No_Symbol'] = df_assets_grouped[f'{column} Var. Abs.']
    
    df_assets_grouped = embed_trend_symbols(df_assets_grouped, ['Banco_Saldo', 'Bienes', 'Inversiones'])
    return df_assets_grouped

def process_income_data(df_income):
    """Process income data with variations and trends."""
    df_income_grouped = df_income.groupby(['Usuario', 'Año Declaración']).agg(
        Ingresos=('Total Ingresos', 'sum'),
        Cant_Ingresos=('Cant_Ingresos', 'sum')
    ).reset_index()

    df_income_grouped = calculate_variation(df_income_grouped, 'Ingresos')
    df_income_grouped[f'Ingresos Var. Abs. No_Symbol'] = df_income_grouped[f'Ingresos Var. Abs.']
    df_income_grouped = embed_trend_symbols(df_income_grouped, ['Ingresos'])
    return df_income_grouped

def calculate_yearly_variations(df):
    """Calculate yearly variations for all columns."""
    df = df.sort_values(['Usuario', 'Año Declaración'])
    
    columns_to_analyze = [
        'Activos', 'Pasivos', 'Patrimonio', 
        'Apalancamiento', 'Endeudamiento',
        'Banco_Saldo', 'Bienes', 'Inversiones', 'Ingresos',
        'Cant_Ingresos'
    ]
    
    # Store original values of absolute changes before formatting
    temp_abs_cols = {}
    
    for column in [col for col in columns_to_analyze if col in df.columns]:
        grouped = df.groupby('Usuario')[column]
        
        for year in [2021, 2022, 2023, 2024]:
            abs_col_name = f'{year} {column} Var. Abs.'
            rel_col_name = f'{year} {column} Var. Rel.'
            
            # Calculate absolute variation (diff)
            abs_variation = grouped.diff()
            df[abs_col_name] = abs_variation
            
            # Store original (unformatted) absolute variation for trend symbol logic
            temp_abs_cols[abs_col_name] = abs_variation
            
            # Calculate relative variation (pct_change)
            pct_change = grouped.pct_change(fill_method=None) * 100
            df[rel_col_name] = pct_change.apply(
                lambda x: f"{x:.2f}%" if pd.notna(x) else ""
            )
            
    # Apply formatting and symbols after all calculations
    for column in [col for col in columns_to_analyze if col in df.columns]:
        for year in [2021, 2022, 2023, 2024]:
            abs_col_name = f'{year} {column} Var. Abs.'
            rel_col_name = f'{year} {column} Var. Rel.'
            
            if abs_col_name in df.columns:
                df[abs_col_name] = df.apply(
                    lambda row: (
                        f"{temp_abs_cols[abs_col_name].loc[row.name]:.2f} {get_trend_symbol(row[rel_col_name])}" 
                        if pd.notna(temp_abs_cols[abs_col_name].loc[row.name]) else "N/A"
                    ),
                    axis=1
                )
            if rel_col_name in df.columns:
                df[rel_col_name] = df.apply(
                    lambda row: (
                        f"{row[rel_col_name]} {get_trend_symbol(row[rel_col_name])}" 
                        if row[rel_col_name] != "" else ""
                    ), 
                    axis=1
                )
    
    return df

def calculate_sudden_wealth_increase(df):
    """Calculate sudden wealth increase rate (Aum. Pat. Subito) as decimal with 1 decimal place"""
    df = df.sort_values(['Usuario', 'Año Declaración'])
    
    # Calculate total wealth (Activo + Patrimonio)
    df['Capital'] = df['Activos'] + df['Patrimonio']
    
    # Calculate year-to-year change as decimal
    df['Aum. Pat. Subito_No_Symbol'] = df.groupby('Usuario')['Capital'].pct_change(fill_method=None)
    
    # Format as decimal (1 place) with trend symbol
    df['Aum. Pat. Subito'] = df['Aum. Pat. Subito_No_Symbol'].apply(
        lambda x: f"{x:.1f} {get_trend_symbol(f'{x*100:.1f}%')}" if pd.notna(x) else "N/A"
    )
    
    return df

def save_results(df, excel_filename="core/src/trends.xlsx"):
    """Save results to Excel with modified column names."""
    try:
        # Create a copy of the dataframe to avoid modifying the original
        df_output = df.copy()
        
        # Convert Usuario to string if it exists (before renaming)
        if 'Usuario' in df_output.columns:
            df_output['Usuario'] = df_output['Usuario'].astype(str)
        
        # Rename columns for output and drop temporary columns
        cols_to_drop = [col for col in df_output.columns if 'No_Symbol' in col]
        df_output = df_output.drop(columns=cols_to_drop, errors='ignore')

        df_output.columns = [col.replace('Usuario', 'Id').replace('Compañía', 'Compania') 
                           for col in df_output.columns]
        
        # Ensure Id is string after renaming
        if 'Id' in df_output.columns:
            df_output['Id'] = df_output['Id'].astype(str)
        
        df_output.to_excel(excel_filename, index=False)
        print(f"Data saved to {excel_filename}")
    except Exception as e:
        print(f"Error saving file: {e}")

def main():
    """Main function to process all data and generate analysis files."""
    try:
        # Process worth data
        df_worth = pd.read_excel("core/src/worthNets.xlsx")
        df_worth = df_worth.rename(columns={
            'Total Activos': 'Activos',
            'Total Pasivos': 'Pasivos',
            'Total Patrimonio': 'Patrimonio'
        })
        
        df_worth = calculate_leverage(df_worth)
        df_worth = calculate_debt_level(df_worth)
        df_worth = calculate_sudden_wealth_increase(df_worth)
        
        for column in ['Activos', 'Pasivos', 'Patrimonio', 'Apalancamiento', 'Endeudamiento']:
            df_worth = calculate_variation(df_worth, column)
            # Create a temporary column to hold the absolute variation without symbol for the N/A check
            df_worth[f'{column} Var. Abs. No_Symbol'] = df_worth[f'{column} Var. Abs.']

        df_worth = embed_trend_symbols(df_worth, ['Activos', 'Pasivos', 'Patrimonio', 'Apalancamiento', 'Endeudamiento'])
        
        # Process asset data
        df_assets = pd.read_excel("core/src/assetNets.xlsx")
        df_assets_processed = process_asset_data(df_assets)
        
        # Process income data
        df_income = pd.read_excel("core/src/incomeNets.xlsx")
        df_income_processed = process_income_data(df_income)
        
        # Merge all data
        df_combined = pd.merge(df_worth, df_assets_processed, on=['Usuario', 'Año Declaración'], how='left')
        df_combined = pd.merge(df_combined, df_income_processed, on=['Usuario', 'Año Declaración'], how='left')
        
        # Calculate yearly variations for the combined dataframe
        df_combined = calculate_yearly_variations(df_combined)

        # Save basic trends
        save_results(df_combined, "core/src/trends.xlsx")
        
    except FileNotFoundError as e:
        print(f"Error: Required file not found - {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
"@

# Create idTrends.py
Set-Content -Path "core/idTrends.py" -Value @"
import pandas as pd
import re
import numpy as np

def get_trend_symbol(value):
    """Determine the trend symbol based on the percentage change."""
    try:
        value_float = float(value.strip('%')) / 100
        if pd.isna(value_float):
            return "➡️"
        elif value_float > 0.1:  # more than 10% increase
            return "📈"
        elif value_float < -0.1:  # more than 10% decrease
            return "📉"
        else:
            return "➡️"  # relatively stable
    except Exception:
        return "➡️"

def clean_and_convert(value, keep_trend=False):
    """Clean and convert to float, optionally preserving trend symbol."""
    if pd.isna(value):
        return value
    
    str_value = str(value)
    
    # Handle "N/A ➡️" case specifically
    if "N/A ➡️" in str_value:
        return np.nan
    
    if keep_trend:
        # Extract numeric part (including percentages)
        numeric_part = re.sub(r'[^\d.%\-]', '', str_value)
        try:
            numeric_value = float(numeric_part.strip('%')) / 100 if '%' in numeric_part else float(numeric_part)
            trend_symbol = get_trend_symbol(str_value)
            return f"{numeric_value:.2%}"[:-1] + trend_symbol  # Format as percentage without % and add symbol
        except:
            return None
    else:
        # For absolute values, just clean numbers
        cleaned = re.sub(r'[^\d.-]', '', str_value)
        try:
            return float(cleaned) if cleaned else None
        except:
            return None

def remove_trend_symbol(value):
    """Remove the trend symbol from a string value."""
    if pd.isna(value):
        return value
    str_value = str(value)
    # Remove any known trend symbols
    cleaned_value = str_value.replace("📈", "").replace("📉", "").replace("➡️", "").strip()
    return cleaned_value


# Read the Excel file
file_path_trends = 'core/src/trends.xlsx'
df_trends = pd.read_excel(file_path_trends)

# Ensure all specified columns exist (create empty ones if they don't)
required_columns = [
    'Id', 'Nombre', 'Compania', 'Cargo', 'fkIdPeriodo', 'Año Declaración', 
    'Año Creación', 'Activos', 'Cant_Bienes', 'Cant_Bancos', 'Cant_Cuentas', 
    'Cant_Inversiones', 'Pasivos', 'Cant_Deudas', 'Patrimonio', 'Apalancamiento', 
    'Endeudamiento', 'Capital', 'Aum. Pat. Subito', 'Activos Var. Abs.', 
    'Activos Var. Rel.', 'Pasivos Var. Abs.', 'Pasivos Var. Rel.', 
    'Patrimonio Var. Abs.', 'Patrimonio Var. Rel.', 'Apalancamiento Var. Abs.', 
    'Apalancamiento Var. Rel.', 'Endeudamiento Var. Abs.', 'Endeudamiento Var. Rel.', 
    'Banco_Saldo', 'Bienes', 'Inversiones', 'Banco_Saldo Var. Abs.', 
    'Banco_Saldo Var. Rel.', 'Bienes Var. Abs.', 'Bienes Var. Rel.', 
    'Inversiones Var. Abs.', 'Inversiones Var. Rel.', 'Ingresos', 
    'Cant_Ingresos', 'Ingresos Var. Abs.', 'Ingresos Var. Rel.'
]

# Add any missing columns with NaN values
for col in required_columns:
    if col not in df_trends.columns:
        df_trends[col] = None

# List of columns to convert to float (absolute variation columns)
float_columns = [
    'Activos Var. Abs.', 
    'Pasivos Var. Abs.', 
    'Patrimonio Var. Abs.', 
    'Apalancamiento Var. Abs.', 
    'Endeudamiento Var. Abs.',  
    'Banco_Saldo Var. Abs.', 
    'Bienes Var. Abs.', 
    'Inversiones Var. Abs.', 
    'Ingresos Var. Abs.'
]

# List of columns to clean infinity values and keep trend symbols
trend_columns = [
    'Apalancamiento', 
    'Endeudamiento', 
    'Activos Var. Rel.', 
    'Pasivos Var. Rel.', 
    'Patrimonio Var. Rel.', 
    'Apalancamiento Var. Rel.', 
    'Endeudamiento Var. Rel.', 
    'Banco_Saldo Var. Rel.', 
    'Bienes Var. Rel.', 
    'Inversiones Var. Rel.', 
    'Ingresos Var. Rel.'
]

# Convert absolute variation columns to float
for col in float_columns:
    if col in df_trends.columns:
        df_trends[col] = df_trends[col].apply(lambda x: clean_and_convert(x, keep_trend=False))

# Process trend columns (handle infinity and preserve trend symbols)
for col in trend_columns:
    if col in df_trends.columns:
        df_trends[col] = df_trends[col].apply(lambda x: clean_and_convert(x, keep_trend=True) 
                          if not pd.isna(x) and str(x).lower() not in ['inf', '-inf', 'inf%'] 
                          else np.nan)

# Special handling for 'Aum. Pat. Subito' column
if 'Aum. Pat. Subito' in df_trends.columns:
    df_trends['Aum. Pat. Subito'] = df_trends['Aum. Pat. Subito'].apply(
        lambda x: np.nan if pd.isna(x) or "N/A ➡️" in str(x) else x
    )

# Reorder columns to match the specified order
df_trends = df_trends[required_columns]

# Read the Personas.xlsx file
file_path_personas = 'core/src/Personas.xlsx'
try:
    df_personas = pd.read_excel(file_path_personas)
except FileNotFoundError:
    print(f"Error: {file_path_personas} not found. Please ensure the file exists.")
    exit()

# You can change 'how' to 'inner', 'right', or 'outer' depending on your desired merge behavior
df_merged = pd.merge(df_trends, df_personas, on='Id', how='left')

# Fill null values in 'Cant_Ingresos' with 0
if 'Cant_Ingresos' in df_merged.columns:
    df_merged['Cant_Ingresos'] = df_merged['Cant_Ingresos'].fillna(0)

# Fill null values in 'Ingresos' with 0
if 'Ingresos' in df_merged.columns:
    df_merged['Ingresos'] = df_merged['Ingresos'].fillna(0)

# Define columns to remove from the output
columns_to_remove = ["Id", "Nombre", "Cargo", "Compania_x", "correo"]

# Drop the specified columns
df_merged = df_merged.drop(columns=columns_to_remove, errors='ignore') # 'errors=ignore' prevents an error if a column isn't found

# Define the desired order of columns
desired_start_columns = ["Cedula", "NOMBRE COMPLETO", "Estado", "Compania_y", "CARGO"]

# Then, get all other columns that are not in the desired_start_columns
remaining_columns = [col for col in df_merged.columns if col not in desired_start_columns]

# Concatenate the two lists to form the final column order
final_column_order = desired_start_columns + remaining_columns

# Reindex the DataFrame with the new column order
df_merged = df_merged[final_column_order]

# Save the modified and merged dataframe back to Excel (idTrends.xlsx)
output_path_idtrends = 'core/src/trendSym.xlsx'
df_merged.to_excel(output_path_idtrends, index=False)
print(f"File has been modified and saved as {output_path_idtrends}")

# --- New section for idTrends.xlsx (without trend symbols) ---
df_idTrends = df_merged.copy() # Create a copy to modify without affecting trendSym.xlsx

# List of columns from which to remove trend symbols (these are the 'Rel.' columns and Apalancamiento/Endeudamiento)
columns_to_clean_symbols = [
    'Apalancamiento', 
    'Endeudamiento', 
    'Activos Var. Rel.', 
    'Pasivos Var. Rel.', 
    'Patrimonio Var. Rel.', 
    'Apalancamiento Var. Rel.', 
    'Endeudamiento Var. Rel.', 
    'Banco_Saldo Var. Rel.', 
    'Bienes Var. Rel.', 
    'Inversiones Var. Rel.', 
    'Ingresos Var. Rel.'
]

for col in columns_to_clean_symbols:
    if col in df_idTrends.columns:
        df_idTrends[col] = df_idTrends[col].apply(remove_trend_symbol)

# Save the dataframe without trend symbols to idTrends.xlsx
output_path_idtrends = 'core/src/idTrends.xlsx'
df_idTrends.to_excel(output_path_idtrends, index=False)
print(f"File without trend symbols has been saved as {output_path_idtrends}")
"@

# Create tcs.py
Set-Content -Path "core/tcs.py" -Value @"
import os
import re
import fitz
import pdfplumber
import pandas as pd
from datetime import datetime

# --- Configuration (can be modified by views.py) ---
categorias_file = "" # Will be set dynamically
cedulas_file = "" # This will now point to Personas.xlsx and be set dynamically
pdf_password = ""


# --- TRM Data (Hardcoded - Pre-calculated Monthly Averages) ---
trm_data = {
    "2024/01/01": 3907.86, "2024/02/01": 3932.79, "2024/03/01": 3902.16,
    "2024/04/01": 3871.93, "2024/05/01": 3866.50, "2024/06/01": 4030.73,
    "2024/07/01": 4040.82, "2024/08/01": 4068.79, "2024/09/01": 4188.08,
    "2024/10/01": 4242.02, "2024/11/01": 4398.81, "2024/12/01": 4381.16,
    "2025/01/01": 4296.84, "2025/02/01": 4125.79, "2025/03/01": 4136.21,
    "2025/04/01": 4272.93, "2025/05/01": 4216.79, "2025/06/01": 4110.15,
    "2025/07/01": 4037.03
}

trm_df = pd.DataFrame()
trm_loaded = False

# Function to load TRM data from the hardcoded dictionary (monthly averages)
def load_trm_data():
    global trm_df, trm_loaded
    try:
        # Convert dictionary to DataFrame
        trm_df = pd.DataFrame(list(trm_data.items()), columns=["Fecha", "TRM"])
        # Convert 'Fecha' column to datetime objects (specifically to the first day of the month)
        trm_df["Fecha"] = pd.to_datetime(trm_df["Fecha"], errors='coerce').dt.date
        trm_loaded = True
        print(f"✅ TRM data loaded successfully from hardcoded monthly average dictionary.")
    except Exception as e:
        print(f"⚠️ Error loading TRM data from dictionary: {e}. MC currency conversion will not be available.")


def obtener_trm(fecha):
    # Ensure fecha is a date object for comparison
    if isinstance(fecha, pd.Timestamp):
        fecha = fecha.date()

    if trm_loaded and pd.isna(fecha):
        return ""
    if trm_loaded:
        # Create a "YYYY/MM/01" string for the given date's month
        lookup_month_start = datetime(fecha.year, fecha.month, 1).date()

        # Find the row in trm_df that matches the start of the month
        fila = trm_df[trm_df["Fecha"] == lookup_month_start]
        if not fila.empty:
            return fila["TRM"].values[0]
    return ""

def formato_excel(valor):
    try:
        if isinstance(valor, (int, float)):
            return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        # Handle cases where value might be a string with ',' for thousands and '.' for decimals (e.g., '1.234,56')
        # or just a string with '.' for thousands and ',' for decimals (e.g., '1,234.56')
        # First, remove thousands separators (both '.' and ',') then replace decimal separator to '.'
        s_valor = str(valor).strip()
        if re.match(r'^\d{1,3}(,\d{3})*(\.\d+)?$', s_valor): # Matches 1,234.56 or 1234.56
            numero = float(s_valor.replace(",", ""))
        elif re.match(r'^\d{1,3}(\.\d{3})*(,\d+)?$', s_valor): # Matches 1.234,56 or 1234,56
            numero = float(s_valor.replace(".", "").replace(",", "."))
        else: # Attempt direct conversion
            numero = float(s_valor)

        return f"{numero:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (ValueError, AttributeError):
        return valor


# --- Categorias Data Loading ---
categorias_df = pd.DataFrame()
categorias_loaded = False

# Modified to accept base_dir
def load_categorias_data(base_dir):
    global categorias_df, categorias_loaded, categorias_file
    categorias_file = os.path.join(base_dir, "core", "src", "categorias.xlsx") # Set the full path here
    if os.path.exists(categorias_file):
        try:
            categorias_df = pd.read_excel(categorias_file)
            if 'Descripción' in categorias_df.columns:
                categorias_df['Descripción'] = categorias_df['Descripción'].astype(str).str.strip()
                categorias_loaded = True
                print(f"✅ Categorias file '{categorias_file}' loaded successfully.")
            else:
                print(f"⚠️ Categorias file '{categorias_file}' loaded, but 'Descripción' column not found. Categorization will not be available.")
        except Exception as e:
            print(f"⚠️ Error loading Categorias file '{categorias_file}': {e}. Categorization will not be available.")
    else:
        print(f"⚠️ Categorias file '{categorias_file}' not found. Categorization will not be available.")

# --- Cedulas Data Loading (now for Personas.xlsx) ---
cedulas_df = pd.DataFrame()
cedulas_loaded = False

# Modified to accept base_dir and reflect new filename/columns
def load_cedulas_data(base_dir):
    global cedulas_df, cedulas_loaded, cedulas_file
    cedulas_file = os.path.join(base_dir, "core", "src", "Personas.xlsx") # Changed filename here
    if os.path.exists(cedulas_file):
        try:
            cedulas_df = pd.read_excel(cedulas_file)
            # Changed column check and conversion to 'NOMBRE COMPLETO' and 'CARGO'
            if 'NOMBRE COMPLETO' in cedulas_df.columns and 'Cedula' in cedulas_df.columns and 'CARGO' in cedulas_df.columns:
                # Apply the clean_cedula_format to the 'Cedula' column upon loading
                cedulas_df['Cedula'] = cedulas_df['Cedula'].apply(clean_cedula_format)
                cedulas_df['NOMBRE COMPLETO'] = cedulas_df['NOMBRE COMPLETO'].astype(str).str.title().str.strip()
                cedulas_loaded = True
                print(f"✅ Personas file '{cedulas_file}' loaded successfully.")
            else:
                print(f"⚠️ Personas file '{cedulas_file}' loaded, but expected columns ('NOMBRE COMPLETO', 'Cedula', 'CARGO') not found. Personas data will not be available.")
        except Exception as e:
            print(f"⚠️ Error loading Personas file '{cedulas_file}': {e}. Personas data will not be available.")
    else:
        print(f"⚠️ Personas file '{cedulas_file}' not found. Personas data will not be available.")

# --- NEW: Function to clean Cedula format (e.g., 123.0 to 123) ---
def clean_cedula_format(value):
    try:
        # If it's a float that represents an integer (e.g., 123.0)
        if isinstance(value, float) and value.is_integer():
            return str(int(value)) # Convert to int, then to string
        # For any other type (string, non-integer float, etc.), convert to string and return as is
        return str(value)
    except (ValueError, TypeError):
        # Handles cases where conversion isn't straightforward (e.g., NaN)
        return str(value) # Ensure it's a string even if it's NaN or an unhandled type


# --- Regex for MC (from mc.py) ---
mc_transaccion_regex = re.compile(
    r"(\w{5,})\s+(\d{2}/\d{2}/\d{4})\s+(.*?)\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)\s+(\d+/\d+)"
)
mc_nombre_regex = re.compile(r"SEÑOR \(A\):\s*(.*)")
mc_tarjeta_regex = re.compile(r"TARJETA:\s+\*{12}(\d{4})")
mc_moneda_regex = re.compile(r"ESTADO DE CUENTA EN:\s+(DOLARES|PESOS)")

# --- Regex for Visa (from visa.py) ---
visa_pattern_transaccion = re.compile(
    r"(\d{6})\s+(\d{2}/\d{2}/\d{4})\s+(.+?)\s+([\d,.]+)\s+([\d,]+)\s+([\d,]+)\s+([\d,.]+)\s+([\d,.]+)\s+(\d+/\d+|0\.00)"
)
visa_pattern_tarjeta = re.compile(r"TARJETA:\s+\*{12}(\d{4})")


# Modified to accept base_dir, input_folder, and output_folder
def run_pdf_processing(base_dir, input_folder, output_folder):
    """
    Main function to process all PDFs in the input_folder.
    This function should be called from views.py.
    """
    global input_base_folder, output_base_folder # Use globals to assign incoming paths
    input_base_folder = input_folder
    output_base_folder = output_folder

    all_resultados = [] # Combined list for all results

    load_trm_data() # This now loads from the hardcoded dictionary (monthly averages)
    load_categorias_data(base_dir) # Pass base_dir
    load_cedulas_data(base_dir) # Pass base_dir (now loading Personas.xlsx)

    if os.path.exists(input_base_folder):
        for archivo in sorted(os.listdir(input_base_folder)):
            if archivo.endswith(".pdf"):
                ruta_pdf = os.path.join(input_base_folder, archivo)

                # Use file name to determine card type
                card_type_is_mc = "MC" in archivo.upper() or "MASTERCARD" in archivo.upper()
                card_type_is_visa = "VISA" in archivo.upper()

                if card_type_is_mc:
                    print(f"📄 Procesando Mastercard: {archivo}")
                    try:
                        with fitz.open(ruta_pdf) as doc:
                            if doc.needs_pass:
                                doc.authenticate(pdf_password)

                            moneda_actual = ""
                            nombre = ""
                            ultimos_digitos = ""
                            tiene_transacciones_mc = False

                            for page_num, page in enumerate(doc, start=1):
                                texto = page.get_text()

                                moneda_match = mc_moneda_regex.search(texto)
                                if moneda_match:
                                    moneda_actual = "USD" if moneda_match.group(1) == "DOLARES" else "COP"

                                if not nombre:
                                    nombre_match = mc_nombre_regex.search(texto)
                                    if nombre_match:
                                        nombre = nombre_match.group(1).strip() # Get raw name

                                if not ultimos_digitos:
                                    tarjeta_match = mc_tarjeta_regex.search(texto)
                                    if tarjeta_match:
                                        ultimos_digitos = tarjeta_match.group(1).strip()

                                for match in mc_transaccion_regex.finditer(texto):
                                    autorizacion, fecha_str, descripcion, valor_original, tasa_pactada, tasa_ea, cargo, saldo, cuotas = match.groups()

                                    if "ABONO DEBITO AUTOMATICO" in descripcion.upper():
                                        continue

                                    try:
                                        fecha_transaccion = pd.to_datetime(fecha_str, dayfirst=True).date()
                                    except:
                                        fecha_transaccion = None

                                    # Pass the date object directly to obtener_trm
                                    tipo_cambio = obtener_trm(fecha_transaccion) if moneda_actual == "USD" else ""

                                    # Determine the value for 'TRM Cierre'
                                    trm_cierre_value = formato_excel(str(tipo_cambio)) if tipo_cambio else "1"

                                    all_resultados.append({
                                        "Archivo": archivo,
                                        "Tipo de Tarjeta": "Mastercard", # New column
                                        "Tarjetahabiente": nombre, # Keep raw name here, convert to title case later for merge
                                        "Número de Tarjeta": ultimos_digitos,
                                        "Moneda": moneda_actual,
                                        "TRM Cierre": trm_cierre_value, # Changed column name and value handling
                                        "Valor Original": formato_excel(valor_original), # Keep this here for calculation later
                                        "Número de Autorización": autorizacion,
                                        "Fecha de Transacción": fecha_transaccion,
                                        "Descripción": descripcion.strip(), # Ensure description is stripped for matching
                                        "Tasa Pactada": formato_excel(tasa_pactada),
                                        "Tasa EA Facturada": formato_excel(tasa_ea),
                                        "Cargos y Abonos": formato_excel(cargo),
                                        "Saldo a Diferir": formato_excel(saldo),
                                        "Cuotas": cuotas,
                                        "Página": page_num,
                                    })
                                    tiene_transacciones_mc = True

                            if not tiene_transacciones_mc and (nombre or ultimos_digitos): # Only add if we found a cardholder/card
                                all_resultados.append({
                                    "Archivo": archivo,
                                    "Tipo de Tarjeta": "Mastercard", # New column
                                    "Tarjetahabiente": nombre, # Keep raw name here, convert to title case later for merge
                                    "Número de Tarjeta": ultimos_digitos,
                                    "Moneda": "",
                                    "TRM Cierre": "1", # Set to 1 for no TRM
                                    "Valor Original": "", # Empty for no transactions
                                    "Número de Autorización": "Sin transacciones",
                                    "Fecha de Transacción": "",
                                    "Descripción": "",
                                    "Tasa Pactada": "",
                                    "Tasa EA Facturada": "",
                                    "Cargos y Abonos": "",
                                    "Saldo a Diferir": "",
                                    "Cuotas": "",
                                    "Página": "",
                                })

                    except Exception as e:
                        print(f"⚠️ Error procesando MC '{archivo}': {e}")

                elif card_type_is_visa:
                    print(f"📄 Procesando Visa: {archivo}")
                    try:
                        with pdfplumber.open(ruta_pdf, password=pdf_password) as pdf:
                            tarjetahabiente_visa = ""
                            tarjeta_visa = ""
                            tiene_transacciones_visa = False
                            last_page_number_visa = 1

                            for page_number, page in enumerate(pdf.pages, start=1):
                                text = page.extract_text()
                                if not text:
                                    continue

                                last_page_number_visa = page_number
                                lines = text.split("\n")

                                for idx, line in enumerate(lines):
                                    line = line.strip()

                                    tarjeta_match_visa = visa_pattern_tarjeta.search(line)
                                    if tarjeta_match_visa:
                                        # Before updating card, if the previous card had no transactions, add a row
                                        if tarjetahabiente_visa and tarjeta_visa and not tiene_transacciones_visa:
                                            all_resultados.append({
                                                "Archivo": archivo,
                                                "Tipo de Tarjeta": "Visa", # New column
                                                "Tarjetahabiente": tarjetahabiente_visa,
                                                "Número de Tarjeta": tarjeta_visa,
                                                "Moneda": "",
                                                "TRM Cierre": "1", # Set to 1 for no TRM
                                                "Valor Original": "", # Empty for no transactions
                                                "Número de Autorización": "Sin transacciones",
                                                "Fecha de Transacción": "",
                                                "Descripción": "",
                                                "Tasa Pactada": "",
                                                "Tasa EA Facturada": "",
                                                "Cargos y Abonos": "",
                                                "Saldo a Diferir": "",
                                                "Cuotas": "",
                                                "Página": last_page_number_visa,
                                            })

                                        tarjeta_visa = tarjeta_match_visa.group(1)
                                        tiene_transacciones_visa = False # Reset for new card

                                        if idx > 0:
                                            posible_nombre = lines[idx - 1].strip()
                                            posible_nombre = (
                                                posible_nombre
                                                .replace("SEÑOR (A):", "")
                                                .replace("Señor (A):", "")
                                                .replace("SEÑOR:", "")
                                                .replace("Señor:", "")
                                                .strip()
                                                #.title() # Not converting here, will convert df column later
                                            )
                                            if len(posible_nombre.split()) >= 2:
                                                tarjetahabiente_visa = posible_nombre
                                        continue

                                    match_visa = visa_pattern_transaccion.search(' '.join(line.split()))
                                    if match_visa and tarjetahabiente_visa and tarjeta_visa:
                                        autorizacion, fecha_str, descripcion, valor_original, tasa_pactada, tasa_ea, cargo, saldo, cuotas = match_visa.groups()

                                        # Visa specific numeric formatting
                                        valor_original_formatted = valor_original.replace(".", "").replace(",", ".")
                                        cargo_formatted = cargo.replace(".", "").replace(",", ".")
                                        saldo_formatted = saldo.replace(".", "").replace(",", ".")

                                        all_resultados.append({
                                            "Archivo": archivo,
                                            "Tipo de Tarjeta": "Visa", # New column
                                            "Tarjetahabiente": tarjetahabiente_visa, # Keep raw name here, convert to title case later for merge
                                            "Número de Tarjeta": tarjeta_visa,
                                            "Moneda": "COP", # Assuming Visa are in COP as no currency explicit extraction
                                            "TRM Cierre": "1", # Not applicable for COP, set to 1
                                            "Valor Original": formato_excel(valor_original_formatted), # Keep this here for calculation later
                                            "Número de Autorización": autorizacion,
                                            "Fecha de Transacción": pd.to_datetime(fecha_str, dayfirst=True).date() if fecha_str else None,
                                            "Descripción": descripcion.strip(), # Ensure description is stripped for matching
                                            "Tasa Pactada": formato_excel(tasa_pactada),
                                            "Tasa EA Facturada": formato_excel(tasa_ea),
                                            "Cargos y Abonos": formato_excel(cargo_formatted),
                                            "Saldo a Diferir": formato_excel(saldo_formatted),
                                            "Cuotas": cuotas,
                                            "Página": page_number,
                                        })
                                        tiene_transacciones_visa = True

                            # After processing all pages for a Visa PDF, check if no transactions were found for the last card processed
                            if tarjetahabiente_visa and tarjeta_visa and not tiene_transacciones_visa:
                                all_resultados.append({
                                    "Archivo": archivo,
                                    "Tipo de Tarjeta": "Visa", # New column
                                    "Tarjetahabiente": tarjetahabiente_visa, # Keep raw name here, convert to title case later for merge
                                    "Número de Tarjeta": tarjeta_visa,
                                    "Moneda": "",
                                    "TRM Cierre": "1", # Set to 1 for no TRM
                                    "Valor Original": "", # Empty for no transactions
                                    "Número de Autorización": "Sin transacciones",
                                    "Fecha de Transacción": "",
                                    "Descripción": "",
                                    "Tasa Pactada": "",
                                    "Tasa EA Facturada": "",
                                    "Cargos y Abonos": "",
                                    "Saldo a Diferir": "",
                                    "Cuotas": "",
                                    "Página": last_page_number_visa,
                                })

                    except Exception as e:
                        print(f"⚠️ Error al procesar Visa '{archivo}': {e}")
                else:
                    print(f"⏩ Archivo '{archivo}' no reconocido como Mastercard o Visa. Saltando.")

    else:
        print(f"⏩ Carpeta de origen '{input_base_folder}' no encontrada. No hay archivos para procesar.")


    # --- Save All Results to a Single Excel File ---
    if all_resultados:
        df_resultado_final = pd.DataFrame(all_resultados)

        # Convert 'Tarjetahabiente' to Title Case for merging with cedulas_df
        # This column will still be named 'Tarjetahabiente' in the dataframe derived from PDFs,
        # but we'll merge it with the 'NOMBRE COMPLETO' from Personas.xlsx
        df_resultado_final['Tarjetahabiente'] = df_resultado_final['Tarjetahabiente'].astype(str).str.title().str.strip()

        # Convert 'Fecha de Transacción' to datetime objects to enable day name extraction
        df_resultado_final['Fecha de Transacción'] = pd.to_datetime(df_resultado_final['Fecha de Transacción'], errors='coerce')

        # Add the 'Día' column
        df_resultado_final['Día'] = df_resultado_final['Fecha de Transacción'].dt.day_name(locale='es_ES').fillna('') # Use 'es_ES' for Spanish day names

        # Add the new 'Tar. x Per.' column
        df_resultado_final['Tar. x Per.'] = df_resultado_final.groupby('Tarjetahabiente')['Número de Tarjeta'].transform('nunique')

        # --- Calculate 'Valor COP' ---
        # Helper function to convert formatted strings to float
        def safe_float_conversion(value):
            try:
                # Remove thousands separator ('.') and replace decimal comma (',') with dot ('.')
                if isinstance(value, str):
                    s_value = value.replace(".", "").replace(",", ".")
                    return float(s_value)
                return float(value)
            except (ValueError, TypeError):
                return pd.NA # Use pandas Not Applicable for missing/invalid values

        df_resultado_final['Valor Original Num'] = df_resultado_final['Valor Original'].apply(safe_float_conversion)
        df_resultado_final['TRM Cierre Num'] = df_resultado_final['TRM Cierre'].apply(safe_float_conversion)

        # Perform the multiplication, handling potential NaNs
        df_resultado_final['Valor COP'] = (df_resultado_final['Valor Original Num'] * df_resultado_final['TRM Cierre Num']).apply(lambda x: formato_excel(x) if pd.notna(x) else '')

        # Drop the temporary numeric columns
        df_resultado_final = df_resultado_final.drop(columns=['Valor Original Num', 'TRM Cierre Num'])


        # Merge with categorias_df if loaded
        if categorias_loaded:
            print("Merging all results with categorias.xlsx...")
            df_resultado_final = pd.merge(df_resultado_final, categorias_df[['Descripción', 'Categoría', 'Subcategoría', 'Zona']],
                                    on='Descripción', how='left')
        else:
            # Add empty columns if categorias.xlsx was not loaded
            df_resultado_final['Categoría'] = ''
            df_resultado_final['Subcategoría'] = ''
            df_resultado_final['Zona'] = ''

        # Merge with cedulas_df (now Personas.xlsx) if loaded
        if cedulas_loaded:
            print("Merging all results with Personas.xlsx...")
            # Ensure the 'Cedula' column in df_resultado_final is also clean before merge
            # This is crucial for accurate merging when Personas.xlsx has already been cleaned.
            if 'Cedula' in cedulas_df.columns: # Check if 'Cedula' exists in the loaded Personas.xlsx
                # Temporarily add a 'Cedula_PDF' column to df_resultado_final to store the cleaned Cedula from Personas.xlsx
                # This ensures we get the *correctly formatted* Cedula from Personas.xlsx during the merge.
                # We'll use 'Tarjetahabiente' as the join key as that's what's derived from PDFs.
                temp_merge_df = pd.merge(
                    df_resultado_final,
                    cedulas_df[['NOMBRE COMPLETO', 'Cedula', 'CARGO', 'AREA']],
                    left_on='Tarjetahabiente',
                    right_on='NOMBRE COMPLETO',
                    how='left',
                    suffixes=('', '_from_personas') # Suffix to avoid column name conflicts
                )

                # Now, transfer the cleaned 'Cedula' from Personas.xlsx to the main DataFrame
                # If a match was found, use the 'Cedula_from_personas', otherwise, keep the existing (or empty) 'Cedula'
                # If 'Cedula' doesn't exist in df_resultado_final yet, create it.
                if 'Cedula' not in df_resultado_final.columns:
                    df_resultado_final['Cedula'] = temp_merge_df['Cedula'].fillna('') # Use Cedula from Personas.xlsx
                else:
                    # If 'Cedula' already exists in df_resultado_final (e.g., from PDF parsing),
                    # prioritize the one from Personas.xlsx if a match was found.
                    df_resultado_final['Cedula'] = temp_merge_df['Cedula'].fillna(df_resultado_final['Cedula'])

                # Transfer 'CARGO' as well
                if 'CARGO' not in df_resultado_final.columns:
                    df_resultado_final['CARGO'] = temp_merge_df['CARGO'].fillna('')
                else:
                    df_resultado_final['CARGO'] = temp_merge_df['CARGO'].fillna(df_resultado_final['CARGO'])

                # Transfer 'AREA' as well
                if 'AREA' not in df_resultado_final.columns:
                    df_resultado_final['AREA'] = temp_merge_df['AREA'].fillna('')
                else:
                    df_resultado_final['AREA'] = temp_merge_df['AREA'].fillna(df_resultado_final['AREA'])

                # Drop the temporary merge column if it was created
                if 'NOMBRE COMPLETO_from_personas' in temp_merge_df.columns:
                    temp_merge_df.drop(columns=['NOMBRE COMPLETO_from_personas'], errors='ignore', inplace=True)
                    
            else:
                print("WARNING: 'Cedula' column not found in Personas.xlsx during merge in tcs.py.")
                df_resultado_final['Cedula'] = ''
                df_resultado_final['CARGO'] = ''
                df_resultado_final['AREA'] = ''

        else:
            # Add empty columns if Personas.xlsx was not loaded
            df_resultado_final['Cedula'] = ''
            df_resultado_final['CARGO'] = '' # Changed from 'Cargo'
            df_resultado_final['AREA'] = ''

        # Define all expected columns in their desired order
        # We will dynamically build this list to ensure all columns exist before selecting them
        # Start with the fixed order for the initial columns
        ordered_columns = [
            "Cedula",
            "Tarjetahabiente",
            "CARGO",
            "AREA", # Add the new AREA column here
            "Tipo de Tarjeta",
            "Número de Tarjeta",
            "Tar. x Per.",
            "Moneda",
            "TRM Cierre",
            "Valor Original",
            "Valor COP",
            "Número de Autorización",
            "Fecha de Transacción",
            "Día",
            "Descripción",
            "Categoría",
            "Subcategoría",
            "Zona",
            "Tasa Pactada",
            "Tasa EA Facturada",
            "Cargos y Abonos",
            "Saldo a Diferir",
            "Cuotas",
            "Página"
            # "Archivo" will be added at the very end
        ]

        # Ensure "Archivo" is added to the end if it exists
        if "Archivo" in df_resultado_final.columns:
            ordered_columns.append("Archivo")

        # Filter to only include columns that actually exist in the DataFrame
        # This loop also ensures the order is maintained based on ordered_columns
        df_resultado_final = df_resultado_final[[col for col in ordered_columns if col in df_resultado_final.columns]]

        # --- IMPORTANT: Apply final Cedula formatting before saving to Excel ---
        # This ensures the 'Cedula' column in the output tcs.xlsx is consistently formatted
        if 'Cedula' in df_resultado_final.columns:
            df_resultado_final['Cedula'] = df_resultado_final['Cedula'].apply(clean_cedula_format)


        # Change the filename to a static name instead of a timestamp
        archivo_salida_unificado = "tcs.xlsx"
        ruta_salida_unificado = os.path.join(output_base_folder, archivo_salida_unificado)
        df_resultado_final.to_excel(ruta_salida_unificado, index=False)
        print(f"\n✅ Archivo unificado de extractos generado correctamente en:\n{ruta_salida_unificado}")
        print("\nPrimeras 5 filas del resultado unificado:")
        print(df_resultado_final.head())
    else:
        print("\n⚠️ No se extrajo ningún dato de los archivos PDF (MC o VISA).")
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
    path('', include('core.urls')), 
    path('accounts/', include('django.contrib.auth.urls')),  
    
]
"@

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

LOGIN_REDIRECT_URL = '/'  
LOGOUT_REDIRECT_URL = '/accounts/login/'  
"@

      # Run migrations
    python3 manage.py makemigrations core
    python3 manage.py migrate
    Write-Host "✅ Database migrations applied successfully." -ForegroundColor $GREEN

    # Start the server
    Write-Host "🚀 Starting Django development server..." -ForegroundColor $GREEN
    python3 manage.py runserver
}

arpa