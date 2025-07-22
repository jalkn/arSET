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
    python -m pip install django whitenoise django-bootstrap-v5 xlsxwriter openpyxl pandas xlrd>=2.0.1 pdfplumber fitz msoffcrypto-tool fuzzywuzzy python-Levenshtein

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
from django.contrib.auth.models import User

class Person(models.Model):
    cedula = models.CharField(max_length=20, primary_key=True)
    nombre_completo = models.CharField(max_length=255)
    correo = models.EmailField(max_length=255, blank=True)
    estado = models.CharField(max_length=50, default='Activo')
    compania = models.CharField(max_length=255, blank=True)
    cargo = models.CharField(max_length=255, blank=True)
    revisar = models.BooleanField(default=False)
    comments = models.TextField(blank=True)
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

class TCS(models.Model):
    id = models.AutoField(primary_key=True)
    person = models.ForeignKey(
        Person,
        on_delete=models.CASCADE,
        related_name='tcs_transactions',
        to_field='cedula',
        db_column='cedula'
    )
    archivo = models.CharField(max_length=255, null=True, blank=True)
    tarjetahabiente = models.CharField(max_length=255, null=True, blank=True)
    numero_tarjeta = models.CharField(max_length=20, null=True, blank=True)
    numero_autorizacion = models.CharField(max_length=50, null=True, blank=True)
    fecha_transaccion = models.DateField(null=True, blank=True)
    descripcion = models.TextField(null=True, blank=True)
    valor_original = models.FloatField(null=True, blank=True)
    tasa_pactada = models.CharField(max_length=50, null=True, blank=True)
    tasa_ea_facturada = models.CharField(max_length=50, null=True, blank=True)
    cargos_abonos = models.FloatField(null=True, blank=True)
    saldo_a_diferir = models.FloatField(null=True, blank=True)
    cuotas = models.CharField(max_length=20, null=True, blank=True)
    pagina = models.IntegerField(null=True, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Transaccion de Tarjeta de Credito"
        verbose_name_plural = "Transacciones de Tarjetas de Credito"
        unique_together = ('person', 'numero_tarjeta', 'fecha_transaccion', 'descripcion', 'valor_original')

    def __str__(self):
        return f"TC Transaction for {self.tarjetahabiente} ({self.numero_tarjeta[-4:]}) on {self.fecha_transaccion}"
"@

# Create admin.py with enhanced configuration
Set-Content -Path "core/admin.py" -Value @" 
from django.contrib import admin
from django import forms
from django.utils.html import format_html
from django.urls import reverse
from core.models import Person, Conflict, FinancialReport, TCS # Import TCS

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
    list_display = ('cedula', 'nombre_completo', 'cargo', 'compania', 'estado', 'revisar')
    search_fields = ('cedula', 'nombre_completo', 'correo')
    list_filter = ('estado', 'compania', 'revisar')
    list_editable = ('revisar',)

    # Custom fields to show in detail view
    readonly_fields = ('cedula_with_actions', 'conflicts_link', 'financial_reports_link', 'tcs_transactions_link')

    fieldsets = (
        (None, {
            'fields': ('cedula_with_actions', 'nombre_completo', 'correo', 'estado', 'compania', 'cargo', 'revisar', 'comments')
        }),
        ('Related Records', {
            'fields': ('conflicts_link', 'financial_reports_link', 'tcs_transactions_link'), # Add tcs_transactions_link
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
                    '<a href="{}" class="changelink">View/Edit Financial Reports</a> &nbsp;'
                    '<a href="{}" class="addlink">Add New Financial Report</a> &nbsp;'
                    '<a href="{}" class="viewlink">All Financial Reports</a>'
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
    financial_reports_link.short_description = 'Financial Reports'
    financial_reports_link.allow_tags = True

    def tcs_transactions_link(self, obj):
        if obj.pk:
            tcs_transaction = obj.tcs_transactions.first()
            if tcs_transaction:
                change_url = reverse('admin:core_tcs_change', args=[tcs_transaction.pk])
                add_url = reverse('admin:core_tcs_add') + f'?person={obj.pk}'
                list_url = reverse('admin:core_tcs_changelist') + f'?q={obj.cedula}'

                return format_html(
                    '<div class="nowrap">'
                    '<a href="{}" class="changelink">View/Edit TCS Transactions</a> &nbsp;'
                    '<a href="{}" class="addlink">Add New TCS Transaction</a> &nbsp;'
                    '<a href="{}" class="viewlink">All TCS Transactions</a>'
                    '</div>',
                    change_url,
                    add_url,
                    list_url
                )
            else:
                add_url = reverse('admin:core_tcs_add') + f'?person={obj.pk}'
                return format_html(
                    '<a href="{}" class="addlink">Create TCS Transaction Record</a>',
                    add_url
                )
        return "-"
    tcs_transactions_link.short_description = 'TCS Transactions'
    tcs_transactions_link.allow_tags = True


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

@admin.register(TCS)
class TCSAdmin(admin.ModelAdmin):
    list_display = (
        'person', 'archivo', 'tarjetahabiente', 'numero_tarjeta',
        'fecha_transaccion', 'descripcion', 'valor_original', 'cargos_abonos'
    )
    search_fields = (
        'person__nombre_completo', 'person__cedula', 'tarjetahabiente',
        'numero_tarjeta', 'descripcion'
    )
    list_filter = ('fecha_transaccion', 'tarjetahabiente')
    raw_id_fields = ('person',)
    fieldsets = (
        (None, {
            'fields': ('person', 'archivo', 'tarjetahabiente', 'numero_tarjeta', 'numero_autorizacion', 'fecha_transaccion')
        }),
        ('Transaction Details', {
            'fields': ('descripcion', 'valor_original', 'tasa_pactada', 'tasa_ea_facturada', 'cargos_abonos', 'saldo_a_diferir', 'cuotas', 'pagina')
        }),
    )
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
from .views import (main, register_superuser, ImportView, person_list,
                   import_conflicts, conflict_list, import_persons, import_tcs,
                   import_finances, person_details, financial_report_list,
                   tcs_list, export_persons_excel, alerts_list) # Import the new view alerts_list

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
    path('import-persons/', views.import_persons, name='import_persons'),
    path('import-conflicts/', views.import_conflicts, name='import_conflicts'),
    path('persons/', views.person_list, name='person_list'),
    path('persons/export/excel/', views.export_persons_excel, name='export_persons_excel'), 
    path('conflicts/', views.conflict_list, name='conflict_list'),
    path('financial-reports/', views.financial_report_list, name='financial_report_list'),
    path('tcs-transactions/', views.tcs_list, name='tcs_list'), 
    path('import-tcs/', views.import_tcs, name='import_tcs'),
    path('import-finances/', views.import_finances, name='import_finances'),
    path('persons/<str:cedula>/', views.person_details, name='person_details'),
    path('alerts/', views.alerts_list, name='alerts_list'),
    path('toggle-revisar/<str:cedula>/', views.toggle_revisar_status, name='toggle_revisar_status'),
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
from django.shortcuts import render
from core.models import Person, Conflict, FinancialReport, TCS 
from django.db.models import Q
import subprocess
import msoffcrypto
import io
import re
from django.views.decorators.http import require_POST
from django.shortcuts import get_object_or_404, redirect

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
        context['conflict_count'] = Conflict.objects.count()
        context['person_count'] = Person.objects.count()
        context['finances_count'] = FinancialReport.objects.count()
        context['tc_count'] = TCS.objects.count() # Get count from TCS model
        context['alerts_count'] = Person.objects.filter(revisar=True).count() # Add alerts count

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

        # --- Status for Tarjetas de Credito ---
        # The TCS model now directly tracks this, but we can still show the latest file status
        visa_dir = os.path.join(core_src_dir, 'visa')
        latest_visa_file = None
        latest_mtime = 0
        if os.path.exists(visa_dir):
            for f_name in os.listdir(visa_dir):
                if f_name.startswith('VISA_') and f_name.endswith('.xlsx'):
                    f_path = os.path.join(visa_dir, f_name)
                    mtime = os.path.getmtime(f_path)
                    if mtime > latest_mtime:
                        latest_mtime = mtime
                        latest_visa_file = f_path

        visa_status = {'filename': 'Tarjetas de Credito', 'records': '-', 'status': 'pending', 'last_updated': None, 'error': None}
        if latest_visa_file:
            try:
                df = pd.read_excel(latest_visa_file)
                visa_status['records'] = len(df)
                visa_status['status'] = 'success'
                visa_status['last_updated'] = datetime.fromtimestamp(latest_mtime)
                # context['tc_count'] is now from the model, no need to update from file here
            except Exception as e:
                visa_status['status'] = 'error'
                visa_status['error'] = f"Error reading file: {str(e)}"
        analysis_results.append(visa_status)

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
            context['financial_report_count'] = idtrends_status['records'] # Update financial report count


        context['analysis_results'] = analysis_results
        return context

@login_required
def main(request):
    context = {}
    context['person_count'] = Person.objects.count()
    context['conflict_count'] = Conflict.objects.count()
    context['finances_count'] = FinancialReport.objects.count()
    context['tc_count'] = TCS.objects.count()
    context['active_person_count'] = Person.objects.filter(estado='Activo').count()
    context['accionista_grupo_count'] = Conflict.objects.filter(q3=True).count()
    context['aum_pat_subito_alert_count'] = FinancialReport.objects.filter(aum_pat_subito__gt=2).count()
    context['alerts_count'] = Person.objects.filter(revisar=True).count()

    # Calculate "Restaurantes" count for the home page
    restaurantes_count = TCS.objects.filter(
        Q(descripcion__icontains='CHEF BURGER MILLA DE O') |
        Q(descripcion__icontains='CREPES Y WAFFLES OVIED') |
        Q(descripcion__icontains='CREPES Y WAFFLES B/GDE') |
        Q(descripcion__icontains='CREPES Y WAFFLES') |
        Q(descripcion__icontains='FIORDI PIZZA PLAZA DEL') |
        Q(descripcion__icontains='FIT CHOICES CALLE 99')
    ).count()
    context['restaurantes_count'] = restaurantes_count # Add to context

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
                'id': 'id', # Assuming 'id' column might exist in the input, or we'll add it later
                'nombre completo': 'nombre_completo',
                # Map the input 'correo_normalizado' directly to a temporary name 'raw_correo'
                'correo_normalizado': 'raw_correo',
                'cedula': 'cedula',
                'estado': 'estado',
                'compania': 'compania',
                'cargo': 'cargo',
                'activo': 'activo', # Assuming 'activo' might be an input column for 'estado'
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
                # Keep '@' symbol and periods, convert to lowercase
                df['correo_to_use'] = df['raw_correo'].str.lower() # MODIFIED LINE
            else:
                df['correo_to_use'] = '' # Initialize if no raw email is present

            # Define the columns for the output Excel file including 'Id', 'Estado', and the new 'correo'
            output_columns = ['Id', 'NOMBRE COMPLETO', 'Cedula', 'Estado', 'Compania', 'CARGO', 'correo']
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
            if 'correo_to_use' in df.columns: # Use the newly created 'correo_to_use' column
                output_columns_df['correo'] = df['correo_to_use']

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
                        'correo': row.get('correo_to_use', ''), # Use 'correo_to_use' for the database
                        'estado': row.get('estado', 'Activo'),
                        'compania': row.get('compania', ''),
                        'cargo': row.get('cargo', ''),
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

            subprocess.run(['python', 'core/conflicts.py'], check=True, cwd=settings.BASE_DIR)

            import pandas as pd
            from core.models import Person, Conflict

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
                    messages.warning(request, f"Skipping financial report for {cedula} due to missing fk_id_periodo.")

            except Exception as e:
                messages.error(request, f"Error processing financial report row {row.get('cedula')}: {str(e)}")
                continue

        messages.success(request, f'Reportes financieros importados exitosamente! {len(df)} registros procesados.')
    except Exception as e:
        messages.error(request, f'Error procesando archivo idTrends.xlsx: {str(e)}')

@login_required
def person_list(request):
    search_query = request.GET.get('q', '')
    status_filter = request.GET.get('status', '')
    cargo_filter = request.GET.get('cargo', '')
    compania_filter = request.GET.get('compania', '')

    order_by = request.GET.get('order_by', 'nombre_completo')
    sort_direction = request.GET.get('sort_direction', 'asc')

    persons = Person.objects.all()

    if search_query:
        persons = persons.filter(
            Q(nombre_completo__icontains=search_query) |
            Q(cedula__icontains=search_query) |
            Q(correo__icontains=search_query))

    if status_filter:
        persons = persons.filter(estado=status_filter)

    if cargo_filter:
        persons = persons.filter(cargo=cargo_filter)

    if compania_filter:
        persons = persons.filter(compania=compania_filter)

    if sort_direction == 'desc':
        order_by = f'-{order_by}'
    persons = persons.order_by(order_by)

    # Convert names to title case for display
    for person in persons:
        person.nombre_completo = person.nombre_completo.title()

    cargos = Person.objects.exclude(cargo='').values_list('cargo', flat=True).distinct().order_by('cargo')
    companias = Person.objects.exclude(compania='').values_list('compania', flat=True).distinct().order_by('compania')

    paginator = Paginator(persons, 25)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    context = {
        'persons': page_obj,
        'page_obj': page_obj,
        'cargos': cargos,
        'companias': companias,
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

    order_by = request.GET.get('order_by', 'nombre_completo')
    sort_direction = request.GET.get('sort_direction', 'asc')

    persons = Person.objects.all()

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
    # --- End dynamic column filtering ---

    if sort_direction == 'desc':
        order_by = f'-{order_by}'
    persons = persons.order_by(order_by).distinct() # Use .distinct() to avoid duplicate persons if related objects cause issues

    # Prepare data for DataFrame
    data = []
    for person in persons:
        # Get the latest financial report for the person, or handle if there are multiple
        # For simplicity, let's assume one relevant financial report or get the first one if multiple exist
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
def tcs_list(request):
    """
    View to display TCS (credit card) transactions.
    """
    search_query = request.GET.get('q', '')
    descripcion_filter = request.GET.get('descripcion', '') # Changed from compania_filter
    numero_tarjeta_filter = request.GET.get('numero_tarjeta', '')
    fecha_transaccion_start = request.GET.get('fecha_transaccion_start', '')
    fecha_transaccion_end = request.GET.get('fecha_transaccion_end', '')
    category_filter = request.GET.get('category_filter', '') # New category filter

    order_by = request.GET.get('order_by', 'fecha_transaccion')
    sort_direction = request.GET.get('sort_direction', 'desc') # Default to descending for dates

    tcs_transactions = TCS.objects.select_related('person').all()

    if search_query:
        tcs_transactions = tcs_transactions.filter(
            Q(person__nombre_completo__icontains=search_query) |
            Q(person__cedula__icontains=search_query) |
            Q(descripcion__icontains=search_query) |
            Q(tarjetahabiente__icontains=search_query)
        )

    if descripcion_filter: # Changed filter application
        tcs_transactions = tcs_transactions.filter(descripcion__icontains=descripcion_filter)

    if numero_tarjeta_filter:
        # Filter by last 4 digits or full number
        tcs_transactions = tcs_transactions.filter(numero_tarjeta__icontains=numero_tarjeta_filter)

    if fecha_transaccion_start:
        try:
            start_date = datetime.strptime(fecha_transaccion_start, '%Y-%m-%d').date()
            tcs_transactions = tcs_transactions.filter(fecha_transaccion__gte=start_date)
        except ValueError:
            messages.error(request, "Formato de fecha de inicio inválido.")

    if fecha_transaccion_end:
        try:
            end_date = datetime.strptime(fecha_transaccion_end, '%Y-%m-%d').date()
            tcs_transactions = tcs_transactions.filter(fecha_transaccion__lte=end_date)
        except ValueError:
            messages.error(request, "Formato de fecha de fin inválido.")

    # Apply category filters
    if category_filter:
        if category_filter == 'restaurantes':
            tcs_transactions = tcs_transactions.filter(
                Q(descripcion__icontains='CHEF BURGER MILLA DE O') |
                Q(descripcion__icontains='CREPES Y WAFFLES OVIED') |
                Q(descripcion__icontains='CREPES Y WAFFLES B/GDE') |
                Q(descripcion__icontains='CREPES Y WAFFLES') |
                Q(descripcion__icontains='FIORDI PIZZA PLAZA DEL') |
                Q(descripcion__icontains='FIT CHOICES CALLE 99')
            )
        elif category_filter == 'suscripciones':
            tcs_transactions = tcs_transactions.filter(
                Q(descripcion__icontains='CHRIS HOLA')
                # Add more values later:
                # | Q(descripcion__icontains='ANOTHER SUBSCRIPTION')
            )
        elif category_filter == 'gastos_diversos':
            tcs_transactions = tcs_transactions.filter(
                Q(descripcion__icontains='ECONOMY PARK RIDE MIA')
                # Add more values later:
                # | Q(descripcion__icontains='OTHER EXPENSE')
            )
        elif category_filter == 'compras':
            tcs_transactions = tcs_transactions.filter(
                Q(descripcion__icontains='CARULLA LAS PALMAS')
                # Add more values later:
                # | Q(descripcion__icontains='ANOTHER PURCHASE')
            )
        elif category_filter == 'gastos_vehiculos':
            tcs_transactions = tcs_transactions.filter(
                Q(descripcion__icontains='CIRCUITO 34 CAR WASH')
                # Add more values later:
                # | Q(descripcion__icontains='ANOTHER PURCHASE')
            )


    if sort_direction == 'desc':
        order_by = f'-{order_by}'
    tcs_transactions = tcs_transactions.order_by(order_by)

    # If you still want to display companies for other purposes, keep this line.
    companias = Person.objects.exclude(compania='').values_list('compania', flat=True).distinct().order_by('compania')

    paginator = Paginator(tcs_transactions, 25)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    context = {
        'tcs_transactions': page_obj,
        'page_obj': page_obj,
        'companias': companias, # Still passing companias even if not directly used in the new filter
        'current_order': order_by.lstrip('-'),
        'current_direction': 'desc' if order_by.startswith('-') else 'asc',
        'all_params': {k: v for k, v in request.GET.items() if k not in ['page', 'order_by', 'sort_direction', 'category_filter']},
        'selected_category': category_filter, # Pass the selected category back to the template
        'alerts_count': Person.objects.filter(revisar=True).count(), # Add alerts count
    }

    return render(request, 'tcs.html', context)


@login_required
def import_tcs(request):
    """View for importing credit card data from PDF files and saving to TCS model"""
    if request.method == 'POST' and request.FILES.getlist('visa_pdf_files'):
        pdf_files = request.FILES.getlist('visa_pdf_files')
        password = request.POST.get('visa_pdf_password', '')

        try:
            # Ensure visa directory exists
            visa_dir = os.path.join(settings.BASE_DIR, 'core', 'src', 'visa')
            os.makedirs(visa_dir, exist_ok=True)

            # Save all PDF files
            for pdf_file in pdf_files:
                dest_path = os.path.join(visa_dir, pdf_file.name)
                with open(dest_path, 'wb+') as destination:
                    for chunk in pdf_file.chunks():
                        destination.write(chunk)

            # Process the PDFs using tcs.py
            subprocess.run(['python', 'core/tcs.py'], check=True, cwd=settings.BASE_DIR)

            # Find the latest generated Excel file by tcs.py
            latest_visa_file = None
            latest_mtime = 0
            for f_name in os.listdir(visa_dir):
                if f_name.startswith('VISA_') and f_name.endswith('.xlsx'):
                    f_path = os.path.join(visa_dir, f_name)
                    mtime = os.path.getmtime(f_path)
                    if mtime > latest_mtime:
                        latest_mtime = mtime
                        latest_visa_file = f_path

            record_count = 0
            if latest_visa_file:
                df = pd.read_excel(latest_visa_file)
                df.columns = [col.lower().replace(' ', '_').replace('.', '').replace('á', 'a').replace('é', 'e').replace('í', 'i').replace('ó', 'o').replace('ú', 'u') for col in df.columns]

                for _, row in df.iterrows():
                    try:
                        # Attempt to extract cedula from 'tarjetahabiente' or other means if available
                        # For now, we'll try to find a person by name, which might not be unique.
                        # A better approach would be to have cedula in the PDF output or a mapping.
                        # For this example, we'll use the 'tarjetahabiente' as a loose identifier.
                        tarjetahabiente_name = row.get('tarjetahabiente', '').strip()
                        # Assuming the tarjetahabiente name in the TCS report might match a person's full name
                        person = None
                        if tarjetahabiente_name:
                            # Try to find a person by full name
                            person = Person.objects.filter(nombre_completo__iexact=tarjetahabiente_name).first()
                            # If not found by full name, try to find by a partial match (e.g., first name, last name)
                            if not person and len(tarjetahabiente_name.split()) > 1:
                                first_name = tarjetahabiente_name.split()[0]
                                last_name = tarjetahabiente_name.split()[-1]
                                person = Person.objects.filter(
                                    Q(nombre_completo__icontains=first_name) &
                                    Q(nombre_completo__icontains=last_name)
                                ).first()

                        if not person:
                            # If no person is found, you might want to create a placeholder Person
                            # or skip this transaction. For now, we'll skip.
                            messages.warning(request, f"Skipping TCS transaction for unknown person: {tarjetahabiente_name}")
                            continue

                        # Clean numeric values for float fields
                        valor_original = _clean_numeric_value(row.get('valor_original'))
                        cargos_abonos = _clean_numeric_value(row.get('cargos_y_abonos'))
                        saldo_a_diferir = _clean_numeric_value(row.get('saldo_a_diferir'))

                        # Convert date string to date object
                        fecha_transaccion_str = row.get('fecha_de_transaccion')
                        fecha_transaccion = None
                        if pd.notna(fecha_transaccion_str):
                            try:
                                fecha_transaccion = datetime.strptime(str(fecha_transaccion_str), '%d/%m/%Y').date()
                            except ValueError:
                                try: # Try another common format if the first fails
                                    fecha_transaccion = datetime.strptime(str(fecha_transaccion_str), '%Y-%m-%d %H:%M:%S').date()
                                except ValueError:
                                    messages.warning(request, f"Could not parse date '{fecha_transaccion_str}' for transaction.")


                        TCS.objects.update_or_create(
                            person=person,
                            numero_tarjeta=str(row.get('numero_de_tarjeta', '')),
                            fecha_transaccion=fecha_transaccion,
                            descripcion=row.get('descripcion', ''),
                            valor_original=valor_original,
                            defaults={
                                'archivo': row.get('archivo', ''),
                                'tarjetahabiente': tarjetahabiente_name,
                                'numero_autorizacion': str(row.get('numero_de_autorizacion', '')),
                                'tasa_pactada': str(row.get('tasa_pactada', '')),
                                'tasa_ea_facturada': str(row.get('tasa_ea_facturada', '')),
                                'cargos_abonos': cargos_abonos,
                                'saldo_a_diferir': saldo_a_diferir,
                                'cuotas': str(row.get('cuotas', '')),
                                'pagina': _clean_numeric_value(row.get('pagina')),
                            }
                        )
                        record_count += 1
                    except Exception as e:
                        messages.error(request, f"Error processing TCS row: {row.to_dict()} - {str(e)}")
                        continue

            messages.success(request, f'Archivos de tarjetas procesados exitosamente! {record_count} transacciones importadas al modelo.')
        except subprocess.CalledProcessError as e:
            messages.error(request, f'Error procesando archivos PDF: {str(e)}')
        except Exception as e:
            messages.error(request, f'Error procesando archivos de tarjetas: {str(e)}')

        return HttpResponseRedirect('/import/')

    return HttpResponseRedirect('/import/')

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
                subprocess.run(['python', 'core/cats.py'], check=True, cwd=settings.BASE_DIR)

                # Run nets.py analysis
                subprocess.run(['python', 'core/nets.py'], check=True, cwd=settings.BASE_DIR)

                # Run trends.py analysis
                subprocess.run(['python', 'core/trends.py'], check=True, cwd=settings.BASE_DIR)

                # Run idTrends.py analysis
                subprocess.run(['python', 'core/idTrends.py'], check=True, cwd=settings.BASE_DIR)

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
    try:
        person = Person.objects.get(cedula=cedula)
        conflicts = Conflict.objects.filter(person=person)
        financial_reports = FinancialReport.objects.filter(person=person).order_by('-ano_declaracion', '-fk_id_periodo') # Fetch financial reports
        tcs_transactions = TCS.objects.filter(person=person).order_by('-fecha_transaccion') # Fetch TCS transactions

        context = {
            'myperson': person,
            'conflicts': conflicts,
            'financial_reports': financial_reports, # Pass financial reports to context
            'tcs_transactions': tcs_transactions, # Pass TCS transactions to context
            'alerts_count': Person.objects.filter(revisar=True).count(), # Add alerts count
        }

        return render(request, 'details.html', context)
    except Person.DoesNotExist:
        messages.error(request, f"Person with ID {cedula} not found")
        return redirect('person_list')


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
import pdfplumber
import pandas as pd
import re
import os
from datetime import datetime
import shutil
import traceback

# Configuration
PDF_FOLDER = os.path.join("core", "src", "visa")
OUTPUT_FOLDER = os.path.join("core", "src", "visa")
COLUMN_NAMES = [
    "Archivo", "Tarjetahabiente", "Número de Tarjeta", "Número de Autorización",
    "Fecha de Transacción", "Descripción", "Valor Original",
    "Tasa Pactada", "Tasa EA Facturada", "Cargos y Abonos",
    "Saldo a Diferir", "Cuotas", "Página"
]

# Patterns
TRANSACTION_PATTERN = re.compile(
    r"(\d{6})\s+(\d{2}/\d{2}/\d{4})\s+(.+?)\s+([\d,.]+)\s+([\d,]+)\s+([\d,]+)\s+([\d,.]+)\s+([\d,.]+)\s+(\d+/\d+|0\.00)"
)
CARD_PATTERN = re.compile(r"TARJETA:\s+\*{12}(\d{4})")

def get_password():
    """Get password from password.txt if it exists"""
    password_file = os.path.join(PDF_FOLDER, "password.txt")
    if os.path.exists(password_file):
        with open(password_file, 'r') as f:
            return f.read().strip()
    return ""

def clean_value(value):
    """Clean and format numeric values"""
    return value.replace(".", "#").replace(",", ".").replace("#", ",")

def process_pdf(pdf_path, password):
    """Process a single PDF file and extract data"""
    data_rows = []
    try:
        with pdfplumber.open(pdf_path, password=password) as pdf:
            cardholder = ""
            card_number = ""
            has_transactions = False
            last_page = 1

            for page_num, page in enumerate(pdf.pages, start=1):
                text = page.extract_text()
                if not text:
                    continue

                last_page = page_num
                lines = text.split("\n")

                for idx, line in enumerate(lines):
                    line = line.strip()

                    # Check for card number change
                    card_match = CARD_PATTERN.search(line)
                    if card_match:
                        if cardholder and card_number and not has_transactions:
                            row = [
                                os.path.basename(pdf_path), cardholder, card_number,
                                "Sin transacciones", "", "", "", "", "", "", "", "", last_page
                            ]
                            data_rows.append(row)

                        card_number = card_match.group(1)
                        has_transactions = False

                        # Get cardholder name from previous line
                        if idx > 0:
                            possible_name = lines[idx - 1].strip()
                            possible_name = (
                                possible_name
                                .replace("SEÑOR (A):", "")
                                .replace("Señor (A):", "")
                                .replace("SEÑOR:", "")
                                .replace("Señor:", "")
                                .strip()
                                .title()
                            )
                            if len(possible_name.split()) >= 2:
                                cardholder = possible_name
                        continue

                    # Check for transactions
                    match = TRANSACTION_PATTERN.search(' '.join(line.split()))
                    if match and cardholder and card_number:
                        row_data = list(match.groups())
                        row_data.insert(0, card_number)
                        row_data.insert(0, cardholder)
                        row_data.insert(0, os.path.basename(pdf_path))

                        # Clean numeric values
                        row_data[6] = clean_value(row_data[6])  # Valor Original
                        row_data[9] = clean_value(row_data[9])  # Cargos y Abonos
                        row_data[10] = clean_value(row_data[10])  # Saldo a Diferir

                        row_data.append(page_num)
                        data_rows.append(row_data)
                        has_transactions = True

            # Add entry if no transactions were found
            if cardholder and card_number and not has_transactions:
                row = [
                    os.path.basename(pdf_path), cardholder, card_number,
                    "Sin transacciones", "", "", "", "", "", "", "", "", last_page
                ]
                data_rows.append(row)

    except Exception as e:
        print(f"⚠ Error processing '{os.path.basename(pdf_path)}': {str(e)}")
        traceback.print_exc()

    return data_rows

def cleanup_files():
    """Clean up temporary files"""
    try:
        # Delete all PDFs in the folder
        for filename in os.listdir(PDF_FOLDER):
            file_path = os.path.join(PDF_FOLDER, filename)
            try:
                if os.path.isfile(file_path) and filename.lower().endswith('.pdf'):
                    os.unlink(file_path)
                elif os.path.isfile(file_path) and filename.lower() == 'password.txt':
                    os.unlink(file_path)
            except Exception as e:
                print(f"⚠ Could not delete {file_path}: {e}")
                
        print("✓ Temporary files cleaned up")
    except Exception as e:
        print(f"⚠ Warning: Could not clean all files: {e}")

def main():
    """Main processing function"""
    # Ensure directories exist
    os.makedirs(PDF_FOLDER, exist_ok=True)
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    
    # Get password
    password = get_password()
    
    # Get PDF files
    pdf_files = [
        f for f in os.listdir(PDF_FOLDER) 
        if f.lower().endswith(".pdf") and os.path.isfile(os.path.join(PDF_FOLDER, f))
    ]
    
    if not pdf_files:
        print("⚠ No PDF files found in the visa folder")
        return
    
    # Process all PDFs
    all_data = []
    for pdf_file in pdf_files:
        pdf_path = os.path.join(PDF_FOLDER, pdf_file)
        print(f"📄 Processing: {pdf_file}")
        all_data.extend(process_pdf(pdf_path, password))
    
    # Export to Excel
    if all_data:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = os.path.join(OUTPUT_FOLDER, f"VISA_{timestamp}.xlsx")
        
        df = pd.DataFrame(all_data, columns=COLUMN_NAMES)
        
        # Convert date column to datetime
        if 'Fecha de Transacción' in df.columns:
            df['Fecha de Transacción'] = pd.to_datetime(
                df['Fecha de Transacción'], 
                dayfirst=True,
                errors='coerce'
            )
        
        # Save to Excel
        writer = pd.ExcelWriter(output_file, engine='openpyxl')
        df.to_excel(writer, index=False)
        
        # Auto-adjust column widths
        worksheet = writer.sheets['Sheet1']
        for column in df.columns:
            column_length = max(df[column].astype(str).map(len).max(), len(column))
            col_idx = df.columns.get_loc(column)
            worksheet.column_dimensions[chr(65 + col_idx)].width = column_length + 2
        
        writer.close()
        
        print(f"\n✓ Excel file generated: {output_file}")
        print(f"Processed {len(df)} transactions")
        
        # Clean up files
        cleanup_files()
    else:
        print("\n⚠ No data extracted from PDFs")

if __name__ == "__main__":
    main()
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

@"
/* Table container styles */
.table-container {
    position: relative;
    overflow: auto;
    max-height: calc(100vh - 300px); /* Adjust this value as needed */
}

/* Make the entire table header sticky */
.table-fixed-header {
    position: sticky;
    top: 0;
    z-index: 10; /* Ensure it stays above the table body */
    background-color: white; /* Fallback background for the header area */
}

/* Apply styles to header cells, but remove individual sticky positioning */
.table-fixed-header th {
    background-color: #f8f9fa; /* Match your table header color */
    /* Remove sticky positioning from individual th elements */
    /* position: sticky; */
    /* top: 0; */
    /* z-index: 20; */
}

/* Add a shadow to the fixed header for visual separation */
.table-fixed-header::after {
    content: '';
    position: absolute;
    left: 0;
    right: 0;
    bottom: -5px;
    height: 5px;
    background: linear-gradient(to bottom, rgba(0,0,0,0.1), transparent);
}

/* Styles for fixed columns */
.table-fixed-column {
    position: sticky;
    right: 0;
    background-color: white;
    z-index: 5;
}

.table-fixed-column::before {
    content: '';
    position: absolute;
    top: 0;
    left: -5px;
    width: 5px;
    height: 100%;
    background: linear-gradient(to right, transparent, rgba(0,0,0,0.1));
}

/* New styles for dynamically frozen columns */
.table-frozen-column {
    position: sticky;
    background-color: white; /* Ensure background is solid when frozen */
    z-index: 6; /* Higher than regular cells but lower than fixed-right column if any */
}

.table-frozen-column::after {
    content: '';
    position: absolute;
    top: 0;
    right: -5px; /* Adjust if shadow is desired on the right */
    width: 5px;
    height: 100%;
    background: linear-gradient(to left, rgba(0,0,0,0.1), transparent);
    pointer-events: none; /* Allows clicks on elements behind the shadow */
}


/* Adjust the z-index for header cells to stay above fixed column */
.table-fixed-header th:last-child {
    z-index: 30;
}

/* Ensure the fixed column stays visible when scrolling */
.table-container {
    overflow: auto;
}

/* Table hover effects */
.table-hover tbody tr:hover {
    background-color: rgba(11, 0, 162, 0.05);
}

/* Style for the freeze button to align it nicely */
.freeze-column-btn {
    margin-right: 5px; /* Space between button and text */
    opacity: 0.5; /* Make it subtle when not active */
}

.freeze-column-btn:hover,
.freeze-column-btn.active {
    opacity: 1; /* More visible when hovered or active */
}
"@ | Out-File -FilePath "core/static/css/freeze.css" -Encoding utf8

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

$jsContent = @"
document.addEventListener('DOMContentLoaded', function() {
    const table = document.querySelector('.table');
    if (!table) return;

    const freezeButtons = document.querySelectorAll('.freeze-column-btn');
    let frozenColumns = JSON.parse(localStorage.getItem('frozenColumns')) || [];

    function applyFrozenColumns() {
        // Clear any existing frozen classes and inline styles
        document.querySelectorAll('.table-frozen-column').forEach(el => {
            el.classList.remove('table-frozen-column');
            el.style.left = ''; // Clear inline style
        });
        document.querySelectorAll('.freeze-column-btn').forEach(btn => {
            btn.classList.remove('active');
        });

        let currentLeft = 0;
        frozenColumns.forEach(colIndex => {
            const cellsInColumn = table.querySelectorAll(``td:nth-child(`$`{colIndex + 1}), th:nth-child(`$`{colIndex + 1})``);
            cellsInColumn.forEach(cell => {
                cell.classList.add('table-frozen-column');
                cell.style.left = ``$`{currentLeft}px``;
            });

            // Mark the corresponding freeze button as active
            const button = document.querySelector(``.freeze-column-btn[data-column-index="`$`{colIndex}"]``);
            if (button) {
                button.classList.add('active');
            }

            // Calculate the width of the frozen column to offset the next one
            // This is a simplified approach, in a real complex table with variable widths,
            // you might need a more robust calculation or a library.
            const headerCell = table.querySelector(``th:nth-child(`$`{colIndex + 1})``);
            if (headerCell) {
                currentLeft += headerCell.offsetWidth;
            }
        });
    }

    freezeButtons.forEach(button => {
        button.addEventListener('click', function() {
            const columnIndex = parseInt(this.dataset.columnIndex);
            const indexInFrozen = frozenColumns.indexOf(columnIndex);

            if (indexInFrozen > -1) {
                // Column is already frozen, unfreeze it
                frozenColumns.splice(indexInFrozen, 1);
            } else {
                // Column is not frozen, freeze it
                frozenColumns.push(columnIndex);
                frozenColumns.sort((a, b) => a - b); // Keep columns ordered by index
            }

            localStorage.setItem('frozenColumns', JSON.stringify(frozenColumns));
            applyFrozenColumns();
        });
    });

    // Apply frozen columns on initial load
    applyFrozenColumns();

    // Re-apply frozen columns on window resize to adjust 'left' positions
    window.addEventListener('resize', applyFrozenColumns);
});
"@
$jsContent | Out-File -FilePath "core/static/js/freeze_columns.js" -Encoding utf8

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
            <a href="{% url 'import' %}" class="btn btn-custom-primary" title="Importar">
                <i class="fas fa-database"></i>
            </a>
            <form method="post" action="{% url 'logout' %}" class="d-inline">
                {% csrf_token %}
                <button type="submit" class="btn btn-custom-primary" title="Cerrar sesion">
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
    <script src="{% static 'js/freeze_columns.js' %}"></script>
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
    <a href="{% url 'person_list' %}" class="btn btn-custom-primary">
        <i class="fas fa-users"></i>
    </a>
    <a href="{% url 'financial_report_list' %}" class="btn btn-custom-primary" title="Bienes y Rentas">
        <i class="fas fa-chart-line" style="color: green;"></i>
    </a>
    <a href="{% url 'tcs_list' %}" class="btn btn-custom-primary" title="Tarjetas">
        <i class="far fa-credit-card" style="color: blue;"></i>
    </a>
    <a href="{% url 'conflict_list' %}" class="btn btn-custom-primary">
        <i class="fas fa-balance-scale" style="color: orange;"></i>
    </a>
    <a href="{% url 'alerts_list' %}" class="btn btn-custom-primary" title="Alertas">
        {% if alerts_count > 0 %}
            <span class="badge bg-danger">{{ alerts_count }}</span>
        {% else %}
            <span class="badge bg-secondary">0</span>
        {% endif %}
        <i class="fas fa-bell" style="color: red;"></i>
    </a>
    <a href="{% url 'import' %}" class="btn btn-custom-primary">
        <i class="fas fa-database"></i>
    </a>
    <form method="post" action="{% url 'logout' %}" class="d-inline">
        {% csrf_token %}
        <button type="submit" class="btn btn-custom-primary" title="Cerrar sesion">
            <i class="fas fa-sign-out-alt"></i>
        </button>
    </form>
</div>
{% endblock %}

{% block content %}
<div class="row mb-4">
    <div class="col-md-3 mb-4">
        <a href="{% url 'person_list' %}" class="card h-100 text-decoration-none text-dark">
            <div class="card-body text-center d-flex flex-column justify-content-center align-items-center">
                <i class="fas fa-users fa-3x text-primary mb-2"></i> {# Adjusted margin-bottom #}
                <h5 class="card-title mb-1">Personas</h5> {# Adjusted margin-bottom #}
                <h2 class="card-text">{{ person_count }}</h2> {# Larger text for count #}
            </div>
        </a>
    </div>

    <div class="col-md-3 mb-4">
        <a href="{% url 'conflict_list' %}" class="card h-100 text-decoration-none text-dark">
            <div class="card-body text-center d-flex flex-column justify-content-center align-items-center">
                <i class="fas fa-balance-scale fa-3x text-warning mb-2"></i>
                <h5 class="card-title mb-1">Conflictos</h5>
                <h2 class="card-text">{{ conflict_count }}</h2>
            </div>
        </a>
    </div>

    <div class="col-md-3 mb-4">
        <a href="{% url 'financial_report_list' %}" class="card h-100 text-decoration-none text-dark">
            <div class="card-body text-center d-flex flex-column justify-content-center align-items-center">
                <i class="fas fa-chart-line fa-3x text-success mb-2"></i>
                <h5 class="card-title mb-1">Bienes y Rentas</h5>
                <h2 class="card-text">{{ finances_count }}</h2>
            </div>
        </a>
    </div>

    <div class="col-md-3 mb-4">
        <a href="{% url 'tcs_list' %}" class="card h-100 text-decoration-none text-dark">
            <div class="card-body text-center d-flex flex-column justify-content-center align-items-center">
                <i class="far fa-credit-card fa-3x text-info mb-2"></i>
                <h5 class="card-title mb-1">Tarjetas de Credito</h5>
                <h2 class="card-text">{{ tc_count }}</h2>
            </div>
        </a>
    </div>

    {# New div for Active Persons #}
    <div class="col-md-3 mb-4">
        <a href="{% url 'person_list' %}?status=Activo" class="card h-100 text-decoration-none text-dark">
            <div class="card-body text-center d-flex flex-column justify-content-center align-items-center">
                <i class="fas fa-user-check fa-3x text-success mb-2"></i> {# Using a different icon, e.g., user-check #}
                <h5 class="card-title mb-1">Personas Activas</h5>
                <h2 class="card-text">{{ active_person_count }}</h2>
            </div>
        </a>
    </div>

    {# New div for Accionista del Grupo #}
    <div class="col-md-3 mb-4">
        <a href="{% url 'conflict_list' %}?column=q3&answer=yes" class="card h-100 text-decoration-none text-dark">
            <div class="card-body text-center d-flex flex-column justify-content-center align-items-center">
                <i class="fas fa-handshake fa-3x text-success mb-2"></i> {# Using a different icon, e.g., handshake #}
                <h5 class="card-title mb-1">Accionista del Grupo</h5>
                <h2 class="card-text">{{ accionista_grupo_count }}</h2>
            </div>
        </a>
    </div>

    {# New div for Aum. Pat. Subito > 2 #}
    <div class="col-md-3 mb-4">
        <a href="{% url 'financial_report_list' %}?column=aum_pat_subito&operator=%3E&value=2" class="card h-100 text-decoration-none text-dark">
            <div class="card-body text-center d-flex flex-column justify-content-center align-items-center">
                <i class="fas fa-arrow-alt-circle-up fa-3x text-danger mb-2"></i>
                <h5 class="card-title mb-1">Indice mayor a 2.0</h5>
                <h2 class="card-text">{{ aum_pat_subito_alert_count }}</h2>
            </div>
        </a>
    </div>

    {# New div for Restaurantes count #}
    <div class="col-md-3 mb-4">
        <a href="{% url 'tcs_list' %}?category_filter=restaurantes" class="card h-100 text-decoration-none text-dark">
            <div class="card-body text-center d-flex flex-column justify-content-center align-items-center">
                <i class="fas fa-utensils fa-3x text-secondary mb-2"></i> {# Icon for restaurants #}
                <h5 class="card-title mb-1">Transacciones en Restaurantes</h5>
                <h2 class="card-text">{{ restaurantes_count }}</h2>
            </div>
        </a>
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
    <a href="/" class="btn btn-custom-primary">
        <i class="fas fa-chart-pie" style="color: rgb(255, 111, 0);"></i>
    </a>
    <a href="{% url 'person_list' %}" class="btn btn-custom-primary">
        <i class="fas fa-users"></i>
    </a>
    <a href="{% url 'financial_report_list' %}" class="btn btn-custom-primary" title="Bienes y Rentas">
        <i class="fas fa-chart-line" style="color: green;"></i>
    </a>
    <a href="{% url 'tcs_list' %}" class="btn btn-custom-primary" title="Tarjetas">
        <i class="far fa-credit-card" style="color: blue;"></i>
    </a>
    <a href="{% url 'conflict_list' %}" class="btn btn-custom-primary">
        <i class="fas fa-balance-scale" style="color: orange;"></i>
    </a>
    <a href="{% url 'alerts_list' %}" class="btn btn-custom-primary" title="Alertas">
        {% if alerts_count > 0 %}
            <span class="badge bg-danger">{{ alerts_count }}</span>
        {% else %}
            <span class="badge bg-secondary">0</span>
        {% endif %}
        <i class="fas fa-bell" style="color: red;"></i>
    </a>
    <form method="post" action="{% url 'logout' %}" class="d-inline">
        {% csrf_token %}
        <button type="submit" class="btn btn-custom-primary" title="Cerrar sesion">
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

<div class="row mb-4">
    <!-- Personas Card -->
    <div class="col-md-3 mb-4">
        <div class="card h-100">
            <div class="card-body">
                <form method="post" enctype="multipart/form-data" action="{% url 'import_persons' %}">
                    {% csrf_token %}
                    <div class="mb-3">
                        <input type="file" class="form-control" id="excel_file" name="excel_file" required>
                        <div class="form-text">El archivo debe incluir las columnas: Id, NOMBRE COMPLETO, CARGO, Cedula, Correo, Compania, Estado</div>
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
            <div class="card-footer">
                <div class="d-flex align-items-center">
                    <span class="badge bg-success">
                        {{ person_count }} Personas Registradas
                    </span>
                </div>
            </div>
        </div>
    </div>

    <!-- Conflictos Card -->
    <div class="col-md-3 mb-4">
        <div class="card h-100">
            <div class="card-body">
                <form method="post" enctype="multipart/form-data" action="{% url 'import_conflicts' %}">
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
            <div class="card-footer">
                <div class="d-flex align-items-center">
                    <span class="badge bg-success">
                        {{ conflict_count }} Declaraciones Registradas
                    </span>
                </div>
            </div>
        </div>
    </div>

    <!-- Bienes y Rentas Card -->
    <div class="col-md-3 mb-4">
        <div class="card h-100">
            <div class="card-body">
                <form method="post" enctype="multipart/form-data" action="{% url 'import_finances' %}">
                    {% csrf_token %}
                    <div class="mb-3">
                        <input type="file" class="form-control" id="finances_file" name="finances_file" required>
                        <div class="form-text">El archivo debe ser dataHistoricaPBI.xlsx</div>
                        <div class="mb-3">
                            <input type="password" class="form-control" id="excel_password" name="excel_password">
                            <div class="form-text">Ingrese la clave si el archivo esta protegido</div>
                        </div>
                    </div>
                    <button type="submit" class="btn btn-custom-primary btn-lg text-start">Importar Bienes y Rentas</button>
                </form>
            </div>
            {% for message in messages %}
                {% if 'import_finances' in message.tags %}
                <div class="card-footer">
                    <div class="alert alert-{{ message.tags }} alert-dismissible fade show mb-0">
                        {{ message }}
                        <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
                    </div>
                </div>
                {% endif %}
            {% endfor %}
            <div class="card-footer">
                <div class="d-flex align-items-center">
                    <span class="badge bg-success">
                        {{ finances_count }} Bienes y Rentas Registradas
                    </span>
                </div>
            </div>
        </div>
    </div>

    <!-- TCs Card -->
    <div class="col-md-3 mb-4">
        <div class="card h-100">
            <div class="card-body">
                <form method="post" enctype="multipart/form-data" action="{% url 'import_tcs' %}">
                    {% csrf_token %}
                    <div class="mb-3">
                        <input type="file" class="form-control" id="visa_pdf_files" name="visa_pdf_files" multiple accept=".pdf" required>
                        <div class="form-text">Seleccione los PDFs de extractos de tarjetas</div>
                        <div class="mb-3">
                            <input type="password" class="form-control" id="visa_pdf_password" name="visa_pdf_password">
                            <div class="form-text">Ingrese la clave si los PDFs estÃ¡n protegidos</div>
                        </div>
                    </div>
                    <button type="submit" class="btn btn-custom-primary btn-lg text-start">Importar TCs</button>
                </form>
            </div>
            {% for message in messages %}
                {% if 'import_tcs' in message.tags %}
                <div class="card-footer">
                    <div class="alert alert-{{ message.tags }} alert-dismissible fade show mb-0">
                        {{ message }}
                        <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
                    </div>
                </div>
                {% endif %}
            {% endfor %}
            <div class="card-footer">
                <div class="d-flex align-items-center">
                    <span class="badge bg-success">
                        {{ tc_count }} Tarjetas Registradas
                    </span>
                </div>
            </div>
        </div>
    </div>

<!-- Analysis Results Row -->
<div class="row">
    <div class="col-12">
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

# Create persons template
@"
{% extends "master.html" %}
{% load static %}

{% block title %}Personas{% endblock %}
{% block navbar_title %}Personas{% endblock %}

{% block navbar_buttons %}
<div>
    <a href="/" class="btn btn-custom-primary">
        <i class="fas fa-chart-pie" style="color: rgb(255, 111, 0);"></i>
    </a>
    <a href="{% url 'financial_report_list' %}" class="btn btn-custom-primary" title="Bienes y Rentas">
        <i class="fas fa-chart-line" style="color: green;"></i>
    </a>
    <a href="{% url 'tcs_list' %}" class="btn btn-custom-primary" title="Tarjetas">
        <i class="far fa-credit-card" style="color: blue;"></i>
    </a>
    <a href="{% url 'conflict_list' %}" class="btn btn-custom-primary">
        <i class="fas fa-balance-scale" style="color: orange;"></i>
    </a>
    <a href="{% url 'alerts_list' %}" class="btn btn-custom-primary" title="Alertas">
        {% if alerts_count > 0 %}
            <span class="badge bg-danger">{{ alerts_count }}</span>
        {% else %}
            <span class="badge bg-secondary">0</span>
        {% endif %}
        <i class="fas fa-bell" style="color: red;"></i>
    </a>
    <a href="{% url 'import' %}" class="btn btn-custom-primary">
        <i class="fas fa-database"></i> 
    </a>
    <a href="{% url 'export_persons_excel' %}{% if request.GET %}?{{ request.GET.urlencode }}{% endif %}" class="btn btn-custom-primary">
        <i class="fas fa-file-excel" style="color: green;"></i>
    </a>
    <form method="post" action="{% url 'logout' %}" class="d-inline">
        {% csrf_token %}
        <button type="submit" class="btn btn-custom-primary" title="Cerrar sesion">
            <i class="fas fa-sign-out-alt"></i>
        </button>
    </form>
</div>
{% endblock %}

{% block content %}
<!-- Search Form -->
<div class="card mb-4 border-0 shadow" style="background-color:rgb(224, 224, 224);">
    <div class="card-body">
        <form method="get" action="." class="row g-3 align-items-center">
            <div class="d-flex align-items-center">
                <span class="badge bg-success">
                    {{ page_obj.paginator.count }} registros
                </span>
                {% if request.GET.q or request.GET.status or request.GET.cargo or request.GET.compania %}
                {% endif %}
            </div>
            <!-- General Search -->
            <div class="col-md-4">
                <input type="text" 
                       name="q" 
                       class="form-control form-control-lg" 
                       placeholder="Buscar persona o cedula" 
                       value="{{ request.GET.q }}">
            </div>
            
            <!-- Status Filter -->
            <div class="col-md-2">
                <select name="status" class="form-select form-select-lg">
                    <option value="">Estado</option>
                    <option value="Activo" {% if request.GET.status == 'Activo' %}selected{% endif %}>Activo</option>
                    <option value="Retirado" {% if request.GET.status == 'Retirado' %}selected{% endif %}>Retirado</option>
                </select>
            </div>
            
            <!-- Cargo Filter -->
            <div class="col-md-2">
                <select name="cargo" class="form-select form-select-lg">
                    <option value="">Cargo</option>
                    {% for cargo in cargos %}
                        <option value="{{ cargo }}" {% if request.GET.cargo == cargo %}selected{% endif %}>{{ cargo }}</option>
                    {% endfor %}
                </select>
            </div>
            
            <!-- Compania Filter -->
            <div class="col-md-2">
                <select name="compania" class="form-select form-select-lg">
                    <option value="">Compania</option>
                    {% for compania in companias %}
                        <option value="{{ compania }}" {% if request.GET.compania == compania %}selected{% endif %}>{{ compania }}</option>
                    {% endfor %}
                </select>
            </div>
            
            <!-- Submit Buttons -->
            <div class="col-md-2 d-flex gap-2">
                <button type="submit" class="btn btn-custom-primary btn-lg flex-grow-1"><i class="fas fa-filter"></i></button>
                <a href="." class="btn btn-custom-primary btn-lg flex-grow-1"><i class="fas fa-undo"></i></a>
            </div>
        </form>
    </div>
</div>

<!-- Persons Table -->
<div class="card border-0 shadow">
    <div class="card-body p-0">
        <div class="table-responsive table-container">
            <table class="table table-striped table-hover mb-0">
                <thead class="table-fixed-header">
                    <tr>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=revisar&sort_direction={% if current_order == 'revisar' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Revisar
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=cedula&sort_direction={% if current_order == 'cedula' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Cedula
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=nombre_completo&sort_direction={% if current_order == 'nombre_completo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Nombre
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=cargo&sort_direction={% if current_order == 'cargo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Cargo
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=correo&sort_direction={% if current_order == 'correo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Correo
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=compania&sort_direction={% if current_order == 'compania' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Compania
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=estado&sort_direction={% if current_order == 'estado' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Estado
                            </a>
                        </th>
                        <th style="color: rgb(0, 0, 0);">Comentarios</th>
                        <th class="table-fixed-column" style="color: rgb(0, 0, 0);">Ver</th>
                    </tr>
                </thead>
                <tbody>
                    {% for person in persons %}
                        <tr {% if person.revisar %}class="table-warning"{% endif %}>
                            <td>
                                <a href="/admin/core/person/{{ person.cedula }}/change/" style="text-decoration: none;" title="{% if person.revisar %}Marcado para revisar{% else %}No marcado{% endif %}">
                                    <i class="fas fa-{% if person.revisar %}check-square text-warning{% else %}square text-secondary{% endif %}" style="padding-left: 20px;"></i>
                                </a>
                            </td>
                            <td>{{ person.cedula }}</td>
                            <td>{{ person.nombre_completo }}</td>
                            <td>{{ person.cargo }}</td>
                            <td>{{ person.correo }}</td>
                            <td>{{ person.compania }}</td>
                            <td>
                                <span class="badge bg-{% if person.estado == 'Activo' %}success{% else %}danger{% endif %}">
                                    {{ person.estado }}
                                </span>
                            </td>
                            <td>{{ person.comments|truncatechars:30|default:"" }}</td>
                            <td class="table-fixed-column">
                                <a href="{% url 'person_details' person.cedula %}" 
                                   class="btn btn-custom-primary btn-sm"
                                   title="View details">
                                    <i class="bi bi-person-vcard-fill"></i>
                                </a>
                            </td>
                        </tr>
                    {% empty %}
                        <tr>
                            <td colspan="9" class="text-center py-4">
                                {% if request.GET.q or request.GET.status or request.GET.cargo or request.GET.compania %}
                                    Sin registros que coincidan con los filtros.
                                {% else %}
                                    Sin registros
                                {% endif %}
                            </td>
                        </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
        
        <!-- Pagination -->
        {% if page_obj.has_other_pages %}
        <div class="p-3">
            <nav aria-label="Page navigation">
                <ul class="pagination justify-content-center">
                    {% if page_obj.has_previous %}
                        <li class="page-item">
                            <a class="page-link" href="?page=1{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="First">
                                <span aria-hidden="true">&laquo;&laquo;</span>
                            </a>
                        </li>
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.previous_page_number }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Previous">
                                <span aria-hidden="true">&laquo;</span>
                            </a>
                        </li>
                    {% endif %}
                    
                    {% for num in page_obj.paginator.page_range %}
                        {% if page_obj.number == num %}
                            <li class="page-item active"><a class="page-link" href="#">{{ num }}</a></li>
                        {% elif num > page_obj.number|add:'-3' and num < page_obj.number|add:'3' %}
                            <li class="page-item"><a class="page-link" href="?page={{ num }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}">{{ num }}</a></li>
                        {% endif %}
                    {% endfor %}
                    
                    {% if page_obj.has_next %}
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.next_page_number }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Next">
                                <span aria-hidden="true">&raquo;</span>
                            </a>
                        </li>
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.paginator.num_pages }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Last">
                                <span aria-hidden="true">&raquo;&raquo;</span>
                            </a>
                        </li>
                    {% endif %}
                </ul>
            </nav>
        </div>
        {% endif %}
    </div>
</div>
{% endblock %}
"@ | Out-File -FilePath "core/templates/persons.html" -Encoding utf8

# conflicts template
@'
{% extends "master.html" %}

{% block title %}Conflictos{% endblock %}
{% block navbar_title %}Conflictos{% endblock %}

{% block navbar_buttons %}
<div>
    <a href="/" class="btn btn-custom-primary">
        <i class="fas fa-chart-pie" style="color: rgb(255, 111, 0);"></i>
    </a>
    <a href="{% url 'person_list' %}" class="btn btn-custom-primary">
        <i class="fas fa-users"></i>
    </a>
    <a href="{% url 'financial_report_list' %}" class="btn btn-custom-primary" title="Bienes y Rentas">
        <i class="fas fa-chart-line" style="color: green;"></i>
    </a>
    <a href="{% url 'tcs_list' %}" class="btn btn-custom-primary" title="Tarjetas">
        <i class="far fa-credit-card" style="color: blue;"></i>
    </a>
    <a href="{% url 'alerts_list' %}" class="btn btn-custom-primary" title="Alertas">
        {% if alerts_count > 0 %}
            <span class="badge bg-danger">{{ alerts_count }}</span>
        {% else %}
            <span class="badge bg-secondary">0</span>
        {% endif %}
        <i class="fas fa-bell" style="color: red;"></i>
    </a>
    <a href="{% url 'import' %}" class="btn btn-custom-primary" title="Importar">
        <i class="fas fa-database"></i>
    </a>
    <form method="post" action="{% url 'logout' %}" class="d-inline">
        {% csrf_token %}
        <button type="submit" class="btn btn-custom-primary" title="Cerrar sesion">
            <i class="fas fa-sign-out-alt"></i>
        </button>
    </form>
</div>
{% endblock %}

{% block content %}
<div class="card mb-4 border-0 shadow" style="background-color:rgb(224, 224, 224);">
    <div class="card-body">
        <form method="get" action="." class="row g-3 align-items-center no-loading">
            <div class="d-flex align-items-center mb-3"> {# Added mb-3 for spacing #}
                <span class="badge bg-success me-2"> {# Added me-2 for spacing #}
                    {{ page_obj.paginator.count }} registros
                </span>
                {% if request.GET.q or request.GET.compania or request.GET.column or request.GET.answer %}
                {% endif %}
            </div>

            <div class="col-md-4">
                <input type="text"
                       name="q"
                       class="form-control form-control-lg"
                       placeholder="Buscar Persona..."
                       {% if request.GET.q %}autofocus{% endif %}
                       value="{{ request.GET.q }}">
            </div>

            <div class="col-md-2">
                <select name="compania" class="form-select form-select-lg">
                    <option value="">Compania</option>
                    {% for compania in companias %}
                        <option value="{{ compania }}" {% if request.GET.compania == compania %}selected{% endif %}>{{ compania }}</option>
                    {% endfor %}
                </select>
            </div>

            <div class="col-md-2">
                <select name="column" class="form-select form-select-lg">
                    <option value="">Selecciona Pregunta</option>
                    <option value="q1" {% if request.GET.column == 'q1' %}selected{% endif %}>Accionista de proveedor</option>
                    <option value="q2" {% if request.GET.column == 'q2' %}selected{% endif %}>Familiar de accionista/empleado</option>
                    <option value="q3" {% if request.GET.column == 'q3' %}selected{% endif %}>Accionista del grupo</option>
                    <option value="q4" {% if request.GET.column == 'q4' %}selected{% endif %}>Actividades extralaborales</option>
                    <option value="q5" {% if request.GET.column == 'q5' %}selected{% endif %}>Negocios con empleados</option>
                    <option value="q6" {% if request.GET.column == 'q6' %}selected{% endif %}>Participacion en juntas</option>
                    <option value="q7" {% if request.GET.column == 'q7' %}selected{% endif %}>Otro conflicto</option>
                    <option value="q8" {% if request.GET.column == 'q8' %}selected{% endif %}>Conoce codigo de conducta</option>
                    <option value="q9" {% if request.GET.column == 'q9' %}selected{% endif %}>Veracidad de informacion</option>
                    <option value="q10" {% if request.GET.column == 'q10' %}selected{% endif %}>Familiar de funcionario</option>
                    <option value="q11" {% if request.GET.column == 'q11' %}selected{% endif %}>Relacion con sector publico</option>
                </select>
            </div>

            <div class="col-md-2">
                <select name="answer" class="form-select form-select-lg">
                    <option value="">Selecciona Respuesta</option>
                    <option value="yes" {% if request.GET.answer == 'yes' %}selected{% endif %}>Si</option>
                    <option value="no" {% if request.GET.answer == 'no' %}selected{% endif %}>No</option>
                    <option value="blank" {% if request.GET.answer == 'blank' %}selected{% endif %}>En Blanco</option> {# Added this line #}
                </select>
            </div>

            <div class="col-md-2 d-flex gap-2">
                <button type="submit" class="btn btn-custom-primary btn-lg flex-grow-1"><i class="fas fa-filter"></i></button>
                <a href="." class="btn btn-custom-primary btn-lg flex-grow-1"><i class="fas fa-undo"></i></a>
            </div>
        </form>
    </div>
</div>

<div class="card border-0 shadow">
    <div class="card-body p-0">
        <div class="table-responsive table-container">
            <table class="table table-striped table-hover mb-0">
                <thead class="table-fixed-header">
                    <tr>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=person__revisar&sort_direction={% if current_order == 'person__revisar' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Revisar
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=person__nombre_completo&sort_direction={% if current_order == 'person__nombre_completo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Nombre
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=person__compania&sort_direction={% if current_order == 'person__compania' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Compania
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=fecha_inicio&sort_direction={% if current_order == 'fecha_inicio' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Ano
                            </a>
                        </th>
                        <th>
                            <i class="fas fa-eye answer-toggle-icon" data-column="answer-q1" style="cursor: pointer;"></i>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=q1&sort_direction={% if current_order == 'q1' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Accionista de proveedor
                            </a>
                        </th>
                        <th class="answer-q1 hidden-answer">Detalle Q1</th>
                        <th>
                            <i class="fas fa-eye answer-toggle-icon" data-column="answer-q2" style="cursor: pointer;"></i>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=q2&sort_direction={% if current_order == 'q2' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Familiar de accionista/empleado
                            </a>
                        </th>
                        <th class="answer-q2 hidden-answer">Detalle Q2</th>
                        <th>
                            <i class="fas fa-eye answer-toggle-icon" data-column="answer-q3" style="cursor: pointer;"></i>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=q3&sort_direction={% if current_order == 'q3' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Accionista del grupo
                            </a>
                        </th>
                        <th class="answer-q3 hidden-answer">Detalle Q3</th>
                        <th>
                            <i class="fas fa-eye answer-toggle-icon" data-column="answer-q4" style="cursor: pointer;"></i>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=q4&sort_direction={% if current_order == 'q4' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Actividades extralaborales
                            </a>
                        </th>
                        <th class="answer-q4 hidden-answer">Detalle Q4</th>
                        <th>
                            <i class="fas fa-eye answer-toggle-icon" data-column="answer-q5" style="cursor: pointer;"></i>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=q5&sort_direction={% if current_order == 'q5' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Negocios con empleados
                            </a>
                        </th>
                        <th class="answer-q5 hidden-answer">Detalle Q5</th>
                        <th>
                            <i class="fas fa-eye answer-toggle-icon" data-column="answer-q6" style="cursor: pointer;"></i>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=q6&sort_direction={% if current_order == 'q6' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Participacion en juntas
                            </a>
                        </th>
                        <th class="answer-q6 hidden-answer">Detalle Q6</th>
                        <th>
                            <i class="fas fa-eye answer-toggle-icon" data-column="answer-q7" style="cursor: pointer;"></i>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=q7&sort_direction={% if current_order == 'q7' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Otro conflicto
                            </a>
                        </th>
                        <th class="answer-q7 hidden-answer">Detalle Q7</th>
                        <th>
                            <i class="fas fa-eye answer-toggle-icon" data-column="answer-q8" style="cursor: pointer;"></i>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=q8&sort_direction={% if current_order == 'q8' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Conoce codigo de conducta
                            </a>
                        </th>
                        <th class="answer-q8 hidden-answer">Detalle Q8</th>
                        <th>
                            <i class="fas fa-eye answer-toggle-icon" data-column="answer-q9" style="cursor: pointer;"></i>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=q9&sort_direction={% if current_order == 'q9' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Veracidad de informacion
                            </a>
                        </th>
                        <th class="answer-q9 hidden-answer">Detalle Q9</th>
                        <th>
                            <i class="fas fa-eye answer-toggle-icon" data-column="answer-q10" style="cursor: pointer;"></i>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=q10&sort_direction={% if current_order == 'q10' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Familiar de funcionario
                            </a>
                        </th>
                        <th class="answer-q10 hidden-answer">Detalle Q10</th>
                        <th>
                            <i class="fas fa-eye answer-toggle-icon" data-column="answer-q11" style="cursor: pointer;"></i>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=q11&sort_direction={% if current_order == 'q11' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Relacion con sector publico
                            </a>
                        </th>
                        <th class="answer-q11 hidden-answer">Detalle Q11</th>
                        <th style="color: rgb(0, 0, 0);">Comentarios</th>
                        <th class="table-fixed-column" style="color: rgb(0, 0, 0);">Ver</th>
                    </tr>
                </thead>
                <tbody>
                    {% for conflict in conflicts %}
                        <tr {% if conflict.person.revisar %}class="table-warning"{% endif %}>
                            <td>
                                <a href="/admin/core/person/{{ conflict.person.cedula }}/change/" style="text-decoration: none;" title="{% if conflict.person.revisar %}Marcado para revisar{% else %}No marcado{% endif %}">
                                    <i class="fas fa-{% if conflict.person.revisar %}check-square text-warning{% else %}square text-secondary{% endif %}" style="padding-left: 20px;"></i>
                                </a>
                            </td>
                            <td>{{ conflict.person.nombre_completo }}</td>
                            <td>{{ conflict.person.compania }}</td>
                            <td>{% if conflict.fecha_inicio %}{{ conflict.fecha_inicio.year }}{% else %}N/A{% endif %}</td> {# Displaying the year #}
                            <td class="text-center">{% if conflict.q1 %}<i style="color: red;">SI</i>{% elif conflict.q1 is False %}<i style="color: green;">NO</i>{% else %}N/A{% endif %}</td>
                            <td class="answer-q1 hidden-answer">{{ conflict.q1_answer|default:"N/A" }}</td> {# Hidden answer column #}
                            <td class="text-center">{% if conflict.q2 %}<i style="color: red;">SI</i>{% elif conflict.q2 is False %}<i style="color: green;">NO</i>{% else %}N/A{% endif %}</td>
                            <td class="answer-q2 hidden-answer">{{ conflict.q2_answer|default:"N/A" }}</td> {# Hidden answer column #}
                            <td class="text-center">{% if conflict.q3 %}<i style="color: red;">SI</i>{% elif conflict.q3 is False %}<i style="color: green;">NO</i>{% else %}N/A{% endif %}</td>
                            <td class="answer-q3 hidden-answer">{{ conflict.q3_answer|default:"N/A" }}</td> {# Hidden answer column #}
                            <td class="text-center">{% if conflict.q4 %}<i style="color: red;">SI</i>{% elif conflict.q4 is False %}<i style="color: green;">NO</i>{% else %}N/A{% endif %}</td>
                            <td class="answer-q4 hidden-answer">{{ conflict.q4_answer|default:"N/A" }}</td> {# Hidden answer column #}
                            <td class="text-center">{% if conflict.q5 %}<i style="color: red;">SI</i>{% elif conflict.q5 is False %}<i style="color: green;">NO</i>{% else %}N/A{% endif %}</td>
                            <td class="answer-q5 hidden-answer">{{ conflict.q5_answer|default:"N/A" }}</td> {# Hidden answer column #}
                            <td class="text-center">{% if conflict.q6 %}<i style="color: red;">SI</i>{% elif conflict.q6 is False %}<i style="color: green;">NO</i>{% else %}N/A{% endif %}</td>
                            <td class="answer-q6 hidden-answer">{{ conflict.q6_answer|default:"N/A" }}</td> {# Hidden answer column #}
                            <td class="text-center">{% if conflict.q7 %}<i style="color: red;">SI</i>{% elif conflict.q7 is False %}<i style="color: green;">NO</i>{% else %}N/A{% endif %}</td>
                            <td class="answer-q7 hidden-answer">{{ conflict.q7_answer|default:"N/A" }}</td> {# Hidden answer column #}
                            <td class="text-center">{% if conflict.q8 %}<i style="color: green;">SI</i>{% elif conflict.q8 is False %}<i style="color: red;">NO</i>{% else %}N/A{% endif %}</td>
                            <td class="answer-q8 hidden-answer">{{ conflict.q8_answer|default:"N/A" }}</td> {# Hidden answer column #}
                            <td class="text-center">{% if conflict.q9 %}<i style="color: green;">SI</i>{% elif conflict.q9 is False %}<i style="color: RED;">NO</i>{% else %}N/A{% endif %}</td>
                            <td class="answer-q9 hidden-answer">{{ conflict.q9_answer|default:"N/A" }}</td> {# Hidden answer column #}
                            <td class="text-center">{% if conflict.q10 %}<i style="color: red;">SI</i>{% elif conflict.q10 is False %}<i style="color: green;">NO</i>{% else %}N/A{% endif %}</td>
                            <td class="answer-q10 hidden-answer">{{ conflict.q10_answer|default:"N/A" }}</td> {# Hidden answer column #}
                            <td class="text-center">{% if conflict.q11 %}<i style="color: red;">SI</i>{% elif conflict.q11 is False %}<i style="color: green;">NO</i>{% else %}N/A{% endif %}</td>
                            <td class="answer-q11 hidden-answer">{{ conflict.q11_answer|default:"N/A" }}</td> {# Hidden answer column #}
                            <td>{{ conflict.person.comments|truncatechars:30|default:"" }}</td>
                            <td class="table-fixed-column">
                                <a href="{% url 'person_details' conflict.person.cedula %}"
                                   class="btn btn-custom-primary btn-sm"
                                   title="View details">
                                    <i class="bi bi-person-vcard-fill"></i>
                                </a>
                            </td>
                        </tr>
                    {% empty %}
                        <tr>
                            <td colspan="36" class="text-center py-4"> {# Adjusted colspan to match the new number of columns #}
                                {% if missing_details_view %} {# Conditional message for the new view #}
                                    No hay conflictos con detalles faltantes para respuestas "SÃƒÆ’Ã‚Â­".
                                {% elif request.GET.q or request.GET.compania or request.GET.column or request.GET.answer %}
                                    Sin registros que coincidan con los filtros.
                                {% else %}
                                    Sin registros de conflictos
                                {% endif %}
                            </td>
                        </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>

        {% if page_obj.has_other_pages %}
        <div class="p-3">
            <nav aria-label="Page navigation">
                <ul class="pagination justify-content-center">
                    {% if page_obj.has_previous %}
                        <li class="page-item">
                            <a class="page-link" href="?page=1{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="First">
                                <span aria-hidden="true">&laquo;&laquo;</span>
                            </a>
                        </li>
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.previous_page_number }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Previous">
                                <span aria-hidden="true">&laquo;</span>
                            </a>
                        </li>
                    {% endif %}

                    {% for num in page_obj.paginator.page_range %}
                        {% if page_obj.number == num %}
                            <li class="page-item active"><a class="page-link" href="#">{{ num }}</a></li>
                        {% elif num > page_obj.number|add:'-3' and num < page_obj.number|add:'3' %}
                            <li class="page-item"><a class="page-link" href="?page={{ num }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}">{{ num }}</a></li>
                        {% endif %}
                    {% endfor %}

                    {% if page_obj.has_next %}
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.next_page_number }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Next">
                                <span aria-hidden="true">&raquo;</span>
                            </a>
                        </li>
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.paginator.num_pages }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Last">
                                <span aria-hidden="true">&raquo;&raquo;</span>
                            </a>
                        </li>
                    {% endif %}
                </ul>
            </nav>
        </div>
        {% endif %}
    </div>
</div>

<style>
    .hidden-answer {
        display: none;
    }
</style>

<script>
    document.addEventListener('DOMContentLoaded', function() {
        // Removed the main toggle button and its listener
        // const toggleButton = document.getElementById('toggleAnswers');

        const answerToggleIcons = document.querySelectorAll('.answer-toggle-icon');

        // Initial state: hide all answer columns and show the info icons
        document.querySelectorAll('.hidden-answer').forEach(cell => {
            cell.style.display = 'none';
        });
        answerToggleIcons.forEach(icon => {
            icon.style.display = 'inline-block'; // Ensure icons are visible
        });


        // Event listener for each individual info icon
        answerToggleIcons.forEach(icon => {
            icon.addEventListener('click', function(event) {
                event.stopPropagation(); // Prevent column sort if clicked on icon
                const columnClass = this.dataset.column; // e.g., "answer-q1"

                // Toggle visibility of all cells (<th> and <td>) with this class
                document.querySelectorAll(`.${columnClass}`).forEach(cell => {
                    if (cell.style.display === 'none' || cell.style.display === '') {
                        cell.style.display = 'table-cell';
                    } else {
                        cell.style.display = 'none';
                    }
                });
            });
        });
    });
</script>
{% endblock %}
'@ | Out-File -FilePath "core/templates/conflicts.html" -Encoding utf8

@"
{% extends "master.html" %}
{% load static %}

{% block title %}Tarjetas{% endblock %}
{% block navbar_title %}Tarjetas{% endblock %}

{% block navbar_buttons %}
<div>
    <a href="/" class="btn btn-custom-primary">
        <i class="fas fa-chart-pie" style="color: rgb(255, 111, 0);"></i>
    </a>
    <a href="{% url 'person_list' %}" class="btn btn-custom-primary">
        <i class="fas fa-users"></i>
    </a>
    <a href="{% url 'financial_report_list' %}" class="btn btn-custom-primary" title="Bienes y Rentas">
        <i class="fas fa-chart-line" style="color: green;"></i>
    </a>
    <a href="{% url 'conflict_list' %}" class="btn btn-custom-primary">
        <i class="fas fa-balance-scale" style="color: orange;"></i>
    </a>
    <a href="{% url 'alerts_list' %}" class="btn btn-custom-primary" title="Alertas">
        {% if alerts_count > 0 %}
            <span class="badge bg-danger">{{ alerts_count }}</span>
        {% else %}
            <span class="badge bg-secondary">0</span>
        {% endif %}
        <i class="fas fa-bell" style="color: red;"></i>
    </a>
    <a href="{% url 'import' %}" class="btn btn-custom-primary">
        <i class="fas fa-database"></i> 
    </a>
    <form method="post" action="{% url 'logout' %}" class="d-inline">
        {% csrf_token %}
        <button type="submit" class="btn btn-custom-primary" title="Cerrar sesion">
            <i class="fas fa-sign-out-alt"></i>
        </button>
    </form>
</div>
{% endblock %}

{% block content %}
<div class="card mb-4 border-0 shadow" style="background-color:rgb(224, 224, 224);">
    <div class="card-body">
        <form method="get" action="." class="row g-3 align-items-center">
            <div class="d-flex align-items-center">
                <span class="badge bg-success">
                    {{ page_obj.paginator.count }} registros
                </span>
                {% if request.GET.q or request.GET.compania or request.GET.numero_tarjeta or request.GET.fecha_transaccion_start or request.GET.fecha_transaccion_end %}
                {% endif %}
            </div>
            <div class="col-md-4">
                <input type="text" 
                       name="q" 
                       class="form-control form-control-lg" 
                       placeholder="Buscar persona, cedula o descripcion..." 
                       value="{{ request.GET.q }}">
            </div>
            
            <div class="col-md-3">
                <input type="text"
                       name="descripcion"
                       class="form-control form-control-lg"
                       placeholder="Buscar por descripcion..."
                       value="{{ request.GET.descripcion }}">
            </div>

            <div class="col-md-3">
                <input type="text" 
                       name="numero_tarjeta" 
                       class="form-control form-control-lg" 
                       placeholder="Numero de Tarjeta (ultimos 4 digitos)" 
                       value="{{ request.GET.numero_tarjeta }}">
            </div>

            <div class="col-md-2 d-flex gap-2">
                <button type="submit" class="btn btn-custom-primary btn-lg flex-grow-1"><i class="fas fa-filter"></i></button>
                <a href="." class="btn btn-custom-primary btn-lg flex-grow-1"><i class="fas fa-undo"></i></a>
            </div>

            <div class="col-12 d-flex flex-wrap gap-2 mt-2">
                <button type="submit" name="category_filter" value="restaurantes"
                        class="btn btn-outline-secondary {% if selected_category == 'restaurantes' %}active{% endif %}">
                    Restaurantes
                </button>
                <button type="submit" name="category_filter" value="suscripciones"
                        class="btn btn-outline-secondary {% if selected_category == 'suscripciones' %}active{% endif %}">
                    Suscripciones
                </button>
                <button type="submit" name="category_filter" value="gastos_diversos"
                        class="btn btn-outline-secondary {% if selected_category == 'gastos_diversos' %}active{% endif %}">
                    Gastos diversos
                </button>
                <button type="submit" name="category_filter" value="compras"
                        class="btn btn-outline-secondary {% if selected_category == 'compras' %}active{% endif %}">
                    Compras
                </button>
                <button type="submit" name="category_filter" value="gastos_vehiculos"
                        class="btn btn-outline-secondary {% if selected_category == 'gastos_vehiculos' %}active{% endif %}">
                    Gastos Vehiculos
                </button>
                <button type="submit" name="category_filter" value="gastos_medicos"
                        class="btn btn-outline-secondary {% if selected_category == 'gastos_medicos' %}active{% endif %}">
                    Gastos Medicos
                </button>
                <button type="submit" name="category_filter" value="tecnologia"
                        class="btn btn-outline-secondary {% if selected_category == 'tecnologia' %}active{% endif %}">
                    Tecnologia
                </button>
                <button type="submit" name="category_filter" value="pagos_online"
                        class="btn btn-outline-secondary {% if selected_category == 'pagos_online' %}active{% endif %}">
                    Pagos online
                </button>
                <button type="submit" name="category_filter" value="telefono_internet"
                        class="btn btn-outline-secondary {% if selected_category == 'telefono_internet' %}active{% endif %}">
                    Servicio telefono e internet
                </button>
                <button type="submit" name="category_filter" value="gastos_viaje"
                        class="btn btn-outline-secondary {% if selected_category == 'gastos_viaje' %}active{% endif %}">
                    Gastos Viaje
                </button>
                <button type="submit" name="category_filter" value="avances"
                        class="btn btn-outline-secondary {% if selected_category == 'avances' %}active{% endif %}">
                    Avances
                </button>
                <button type="submit" name="category_filter" value="gastos_legales"
                        class="btn btn-outline-secondary {% if selected_category == 'gastos_legales' %}active{% endif %}">
                    Gastos Legales
                </button>
            </div>
        </form>
    </div>
</div>

<div class="card border-0 shadow">
    <div class="card-body p-0">
        <div class="table-responsive table-container">
            <table class="table table-striped table-hover mb-0">
                <thead class="table-fixed-header">
                    <tr>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=person__nombre_completo&sort_direction={% if current_order == 'person__nombre_completo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Persona
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=person__cedula&sort_direction={% if current_order == 'person__cedula' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Cedula
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=tarjetahabiente&sort_direction={% if current_order == 'tarjetahabiente' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Tarjetahabiente
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=numero_tarjeta&sort_direction={% if current_order == 'numero_tarjeta' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                No. Tarjeta
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=fecha_transaccion&sort_direction={% if current_order == 'fecha_transaccion' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Fecha Transaccion
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=descripcion&sort_direction={% if current_order == 'descripcion' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Descripcion
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=valor_original&sort_direction={% if current_order == 'valor_original' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Valor Original
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=cargos_abonos&sort_direction={% if current_order == 'cargos_abonos' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Cargos/Abonos
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=saldo_a_diferir&sort_direction={% if current_order == 'saldo_a_diferir' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Saldo a Diferir
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=cuotas&sort_direction={% if current_order == 'cuotas' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Cuotas
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=tasa_pactada&sort_direction={% if current_order == 'tasa_pactada' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Tasa Pactada
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=tasa_ea_facturada&sort_direction={% if current_order == 'tasa_ea_facturada' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Tasa EA Facturada
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=numero_autorizacion&sort_direction={% if current_order == 'numero_autorizacion' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                No. Autorizacion
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=pagina&sort_direction={% if current_order == 'pagina' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Pagina
                            </a>
                        </th>
                        <th class="table-fixed-column" style="color: rgb(0, 0, 0);">Ver</th>
                    </tr>
                </thead>
                <tbody>
                    {% for transaction in tcs_transactions %}
                        <tr>
                            <td>{{ transaction.person.nombre_completo }}</td>
                            <td>{{ transaction.person.cedula }}</td>
                            <td>{{ transaction.tarjetahabiente|default:"N/A" }}</td>
                            <td>{{ transaction.numero_tarjeta|default:"N/A" }}</td>
                            <td>{{ transaction.fecha_transaccion|date:"Y-m-d"|default:"N/A" }}</td>
                            <td>{{ transaction.descripcion|default:"N/A" }}</td>
                            <td>{{ transaction.valor_original|default:"N/A"|floatformat:2 }}</td>
                            <td>{{ transaction.cargos_abonos|default:"N/A"|floatformat:2 }}</td>
                            <td>{{ transaction.saldo_a_diferir|default:"N/A"|floatformat:2 }}</td>
                            <td>{{ transaction.cuotas|default:"N/A" }}</td>
                            <td>{{ transaction.tasa_pactada|default:"N/A" }}</td>
                            <td>{{ transaction.tasa_ea_facturada|default:"N/A" }}</td>
                            <td>{{ transaction.numero_autorizacion|default:"N/A" }}</td>
                            <td>{{ transaction.pagina|default:"N/A" }}</td>
                            <td class="table-fixed-column">
                                <a href="{% url 'person_details' transaction.person.cedula %}" 
                                   class="btn btn-custom-primary btn-sm"
                                   title="View person details">
                                    <i class="bi bi-person-vcard-fill"></i>
                                </a>
                            </td>
                        </tr>
                    {% empty %}
                        <tr>
                            <td colspan="15" class="text-center py-4">
                                {% if request.GET.q or request.GET.compania or request.GET.numero_tarjeta or request.GET.fecha_transaccion_start or request.GET.fecha_transaccion_end or request.GET.category_filter %}
                                    Sin transacciones de tarjetas de credito que coincidan con los filtros.
                                {% else %}
                                    Sin transacciones de tarjetas de credito
                                {% endif %}
                            </td>
                        </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
        
        {% if page_obj.has_other_pages %}
        <div class="p-3">
            <nav aria-label="Page navigation">
                <ul class="pagination justify-content-center">
                    {% if page_obj.has_previous %}
                        <li class="page-item">
                            <a class="page-link" href="?page=1{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="First">
                                <span aria-hidden="true">&laquo;&laquo;</span>
                            </a>
                        </li>
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.previous_page_number }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Previous">
                                <span aria-hidden="true">&laquo;</span>
                            </a>
                        </li>
                    {% endif %}
                    
                    {% for num in page_obj.paginator.page_range %}
                        {% if page_obj.number == num %}
                            <li class="page-item active"><a class="page-link" href="#">{{ num }}</a></li>
                        {% elif num > page_obj.number|add:'-3' and num < page_obj.number|add:'3' %}
                            <li class="page-item"><a class="page-link" href="?page={{ num }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}">{{ num }}</a></li>
                        {% endif %}
                    {% endfor %}
                    
                    {% if page_obj.has_next %}
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.next_page_number }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Next">
                                <span aria-hidden="true">&raquo;</span>
                            </a>
                        </li>
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.paginator.num_pages }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Last">
                                <span aria-hidden="true">&raquo;&raquo;</span>
                            </a>
                        </li>
                    {% endif %}
                </ul>
            </nav>
        </div>
        {% endif %}
    </div>
</div>
{% endblock %}
"@ | Out-File -FilePath "core/templates/tcs.html" -Encoding utf8

# finances template
@" 
{% extends "master.html" %}
{% load static %}
{% load humanize %}

{% block title %}Bienes y Rentas{% endblock %}
{% block navbar_title %}Bienes y Rentas{% endblock %}

{% block navbar_buttons %}
<div>
    <a href="/" class="btn btn-custom-primary">
        <i class="fas fa-chart-pie" style="color: rgb(255, 111, 0);"></i>
    </a>
    <a href="{% url 'person_list' %}" class="btn btn-custom-primary">
        <i class="fas fa-users"></i>
    </a>
    <a href="{% url 'tcs_list' %}" class="btn btn-custom-primary" title="Tarjetas">
        <i class="far fa-credit-card" style="color: blue;"></i>
    </a>
    <a href="{% url 'conflict_list' %}" class="btn btn-custom-primary">
        <i class="fas fa-balance-scale" style="color: orange;"></i>
    </a>
    <a href="{% url 'alerts_list' %}" class="btn btn-custom-primary" title="Alertas">
        {% if alerts_count > 0 %}
            <span class="badge bg-danger">{{ alerts_count }}</span>
        {% else %}
            <span class="badge bg-secondary">0</span>
        {% endif %}
        <i class="fas fa-bell" style="color: red;"></i>
    </a>
    <a href="{% url 'import' %}" class="btn btn-custom-primary">
        <i class="fas fa-database"></i> 
    </a>
    <a href="{% url 'export_persons_excel' %}{% if request.GET %}?{{ request.GET.urlencode }}{% endif %}" class="btn btn-custom-primary">
        <i class="fas fa-file-excel" style="color: green;"></i> 
    </a>
    <form method="post" action="{% url 'logout' %}" class="d-inline">
        {% csrf_token %}
        <button type="submit" class="btn btn-custom-primary" title="Cerrar sesion">
            <i class="fas fa-sign-out-alt"></i>
        </button>
    </form>
</div>
{% endblock %}

{% block content %}
<div class="card mb-4 border-0 shadow" style="background-color:rgb(224, 224, 224);">
    <div class="card-body">
        <form method="get" action="." class="row g-3 align-items-start"> {# Changed align-items-center to align-items-start #}

            <div class="d-flex align-items-center mb-3 col-12"> {# Added col-12 #}
                <span class="badge bg-success me-2"> {# Added me-2 for margin #}
                    {{ page_obj.paginator.count }} registros
                </span>
                {% comment %} Removed the if block here as it was empty {% endcomment %}
            </div>

            <div id="filter-rows-container" class="col-md-9 row g-3">
                <div class="col-md-3">
                    <input type="text" 
                        name="q" 
                        class="form-control form-control-lg" 
                        placeholder="Buscar persona..." 
                        value="{{ request.GET.q|default:'' }}">
                </div>

                <div class="filter-group-template" style="display: none;">
                    <div class="filter-group row g-3 align-items-center">
                        <div class="col-md-4">
                            <select name="column_X" class="form-select form-select-lg column-select">
                                <option value="">Selecciona Columna</option>
                                <option value="fk_id_periodo">Periodo ID</option>
                                <option value="ano_declaracion">Ano Declaracion</option>
                                <option value="activos">Activos</option>
                                <option value="cant_bienes">Cantidad Bienes</option>
                                <option value="cant_bancos">Cantidad Bancos</option>
                                <option value="cant_cuentas">Cantidad Cuentas</option>
                                <option value="cant_inversiones">Cantidad Inversiones</option>
                                <option value="pasivos">Pasivos</option>
                                <option value="cant_deudas">Cantidad Deudas</option>
                                <option value="patrimonio">Patrimonio</option>
                                <option value="endeudamiento">Endeudamiento</option>
                                <option value="aum_pat_subito">Aumento Patrimonio Subito</option>
                                <option value="activos_var_abs">Activos Var. Absoluta</option>
                                <option value="activos_var_rel">Activos Var. Relativa</option>
                                <option value="pasivos_var_abs">Pasivos Var. Absoluta</option>
                                <option value="pasivos_var_rel">Pasivos Var. Relativa</option>
                                <option value="patrimonio_var_abs">Patrimonio Var. Absoluta</option>
                                <option value="patrimonio_var_rel">Patrimonio Var. Relativa</option>
                                <option value="bienes">Bienes</option>
                                <option value="inversiones">Inversiones</option>
                                <option value="banco_saldo_var_abs">Banco Saldo Var. Absoluta</option>
                                <option value="banco_saldo_var_rel">Banco Saldo Var. Relativa</option>
                                <option value="bienes_var_abs">Bienes Var. Absoluta</option>
                                <option value="bienes_var_rel">Bienes Var. Relativa</option>
                                <option value="inversiones_var_abs">Inversiones Var. Absoluta</option>
                                <option value="inversiones_var_rel">Inversiones Var. Relativa</option>
                                <option value="ingresos">Ingresos</option>
                                <option value="cant_ingresos">Cantidad Ingresos</option>
                                <option value="ingresos_var_abs">Ingresos Var. Absoluta</option>
                                <option value="ingresos_var_rel">Ingresos Var. Relativa</option>
                            </select>
                        </div>
                        
                        <div class="col-md-3">
                            <select name="operator_X" class="form-select form-select-lg operator-select">
                                <option value="">Selecciona operador</option>
                                <option value=">">Mayor que</option>
                                <option value="<">Menor que</option>
                                <option value="=">Igual a</option>
                                <option value=">=">Mayor o igual</option>
                                <option value="<=">Menor o igual</option>
                                <option value="between">Entre</option>
                                <option value="contains">Contiene</option>
                            </select>
                        </div>
                        
                        <div class="col-md-3 value1-container">
                            <input type="text"
                                name="value_X"
                                class="form-control form-control-lg value1-input"
                                placeholder="Valor">
                        </div>

                        <div class="col-md-2 value2-container" style="display: none;">
                            <input type="text"
                                name="value2_X"
                                class="form-control form-control-lg value2-input"
                                placeholder="Segundo Valor">
                        </div>
                        <div class="col-md-1 d-flex justify-content-center"> {# Added d-flex justify-content-center for button alignment #}
                            <button type="button" class="btn btn-danger remove-filter-btn"><i class="fas fa-minus"></i></button>
                        </div>
                    </div>
                </div>
            </div> 
                <div class="col-12 d-flex gap-2 justify-content-end mt-3"> {# Added mt-3 for top margin #}
                    <button type="button" class="btn btn-custom-primary btn-lg" id="add-filter-btn"><i class="fas fa-plus"></i></button>
                    <button type="submit" class="btn btn-custom-primary btn-lg"><i class="fas fa-filter"></i></button>
                    <a href="." class="btn btn-custom-primary btn-lg"><i class="fas fa-undo"></i></a>
                </div>
    </div>
</div>

<div class="card border-0 shadow">
    <div class="card-body p-0">
        <div class="table-responsive table-container">
            <table class="table table-striped table-hover mb-0">
                <thead class="table-fixed-header">
                    <tr>
                        <th data-column-index="0"></th>
                        <th data-column-index="1"></th>
                        <th data-column-index="2"></th>
                        <th data-column-index="3"></th>
                        <th data-column-index="4"></th>
                        <th data-column-index="5"></th>
                        <th data-column-index="6"></th>
                        <th data-column-index="7"></th>
                        <th data-column-index="8"></th> 
                        <th data-column-index="9"></th>
                        <th style="background-color: red; color: white;" data-column-index="10">-50%</th>
                        <th data-column-index="11"></th>
                        <th style="background-color: red; color: white;" data-column-index="12">-50%</th>
                        <th data-column-index="13"></th>
                        <th data-column-index="14"></th>
                        <th style="background-color: red; color: white;" data-column-index="15">-50%</th>
                        <th data-column-index="16"></th>
                        <th data-column-index="17"></th>
                        <th data-column-index="18"></th>
                        <th data-column-index="19"></th>
                        <th data-column-index="20"></th>
                        <th data-column-index="21"></th>
                        <th data-column-index="22"></th>
                        <th data-column-index="23"></th>
                        <th data-column-index="24"></th>
                        <th data-column-index="25"></th>
                        <th data-column-index="26"></th>
                        <th data-column-index="27"></th>
                        <th data-column-index="28"></th>
                        <th data-column-index="29"></th>
                        <th data-column-index="30"></th>
                        <th data-column-index="31"></th>
                        <th data-column-index="32"></th>
                        <th data-column-index="33"></th>
                        <th data-column-index="34"></th>
                        <th class="table-fixed-column" data-column-index="35"></th>
                    </tr>
                    <tr>
                        <th data-column-index="0"></th>
                        <th data-column-index="1"></th>
                        <th data-column-index="2"></th>
                        <th data-column-index="3"></th>
                        <th data-column-index="4"></th>
                        <th data-column-index="5"></th>
                        <th data-column-index="6"></th>
                        <th data-column-index="7"></th>
                        <th data-column-index="8"></th> 
                        <th data-column-index="9"></th>
                        <th style="background-color: green; color: white;" data-column-index="10">-30%</th>
                        <th data-column-index="11"></th>
                        <th style="background-color: green; color: white;" data-column-index="12">-30%</th>
                        <th data-column-index="13"></th>
                        <th data-column-index="14"></th>
                        <th style="background-color: green; color: white;" data-column-index="15">-30%</th>
                        <th data-column-index="16"></th>
                        <th data-column-index="17"></th>
                        <th data-column-index="18"></th>
                        <th data-column-index="19"></th>
                        <th data-column-index="20"></th>
                        <th data-column-index="21"></th>
                        <th data-column-index="22"></th>
                        <th data-column-index="23"></th>
                        <th data-column-index="24"></th>
                        <th data-column-index="25"></th>
                        <th data-column-index="26"></th>
                        <th data-column-index="27"></th>
                        <th data-column-index="28"></th>
                        <th data-column-index="29"></th>
                        <th data-column-index="30"></th>
                        <th data-column-index="31"></th>
                        <th data-column-index="32"></th>
                        <th data-column-index="33"></th>
                        <th data-column-index="34"></th>
                        <th class="table-fixed-column" data-column-index="35"></th>
                    </tr>
                    <tr>
                        <th data-column-index="0"></th>
                        <th data-column-index="1"></th>
                        <th data-column-index="2"></th>
                        <th data-column-index="3">Medio</th>
                        <th data-column-index="4">">="</th>
                        <th style="background-color: green; " data-column-index="5"></th>
                        <th data-column-index="6"></th>
                        <th style="background-color: green;  color: white;" data-column-index="7">1.5</th>
                        <th style="background-color: green;  color: white;" data-column-index="8">50%</th> 
                        <th data-column-index="9"></th>
                        <th style="background-color: green; color: white;" data-column-index="10">30%</th>
                        <th data-column-index="11"></th>
                        <th style="background-color: green; color: white;" data-column-index="12">30%</th>
                        <th data-column-index="13"></th>
                        <th data-column-index="14"></th>
                        <th style="background-color: green; color: white;" data-column-index="15">30%</th>
                        <th data-column-index="16"></th>
                        <th style="background-color: green; color: white;" data-column-index="17">4</th>
                        <th data-column-index="18"></th>
                        <th data-column-index="19"></th>
                        <th data-column-index="20"></th>
                        <th data-column-index="21"></th>
                        <th data-column-index="22"></th>
                        <th data-column-index="23"></th>
                        <th data-column-index="24"></th>
                        <th style="background-color: green; color: white;" data-column-index="25">4</th>
                        <th data-column-index="26"></th>
                        <th data-column-index="27"></th>
                        <th data-column-index="28"></th>
                        <th data-column-index="29"></th>
                        <th data-column-index="30"></th>
                        <th data-column-index="31"></th>
                        <th data-column-index="32"></th>
                        <th data-column-index="33"></th>
                        <th data-column-index="34"></th>
                        <th class="table-fixed-column" data-column-index="35"></th>
                    </tr>
                    <tr>
                        <th data-column-index="0"></th>
                        <th data-column-index="1"></th>
                        <th data-column-index="2"></th>
                        <th data-column-index="3">Alto</th>
                        <th data-column-index="4">">"</th>
                        <th style="background-color: red;" data-column-index="5"></th>
                        <th data-column-index="6"></th>
                        <th style="background-color: red; color: white;" data-column-index="7">2</th>
                        <th style="background-color: red; color: white;" data-column-index="8">70%</th> 
                        <th data-column-index="9"></th>
                        <th style="background-color: red; color: white;" data-column-index="10">50%</th>
                        <th data-column-index="11"></th>
                        <th style="background-color: red; color: white;" data-column-index="12">50%</th>
                        <th data-column-index="13"></th>
                        <th data-column-index="14"></th>
                        <th style="background-color: red; color: white;" data-column-index="15">50%</th>
                        <th data-column-index="16"></th>
                        <th style="background-color: red; color: white;" data-column-index="17">6</th>
                        <th data-column-index="18"></th>
                        <th data-column-index="19"></th>
                        <th data-column-index="20"></th>
                        <th data-column-index="21"></th>
                        <th data-column-index="22"></th>
                        <th data-column-index="23"></th>
                        <th data-column-index="24"></th>
                        <th style="background-color: red; color: white;" data-column-index="25">6</th>
                        <th data-column-index="26"></th>
                        <th data-column-index="27"></th>
                        <th data-column-index="28"></th>
                        <th data-column-index="29"></th>
                        <th data-column-index="30"></th>
                        <th data-column-index="31"></th>
                        <th data-column-index="32"></th>
                        <th data-column-index="33"></th>
                        <th data-column-index="34"></th>
                        <th class="table-fixed-column" data-column-index="35"></th>
                    </tr>
                    <tr>
                        <th data-column-index="0">
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="0" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=person__revisar&sort_direction={% if current_order == 'person__revisar' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Revisar
                            </a>
                        </th>
                        <th data-column-index="1">
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="1" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=person__nombre_completo&sort_direction={% if current_order == 'person__nombre_completo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Nombre
                            </a>
                        </th>
                        <th data-column-index="2">
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="2" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=person__compania&sort_direction={% if current_order == 'person__compania' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Compania
                            </a>
                        </th>
                        <th data-column-index="3">
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="3" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=person__cargo&sort_direction={% if current_order == 'person__cargo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Cargo
                            </a>
                        </th>
                        <th style="color: rgb(0, 0, 0);" data-column-index="4">
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="4" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            Comentarios
                        </th>
                        <th data-column-index="5">
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="5" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=fk_id_periodo&sort_direction={% if current_order == 'fk_id_periodo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Periodo
                            </a>
                        </th>
                        <th data-column-index="6">
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="6" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=ano_declaracion&sort_direction={% if current_order == 'ano_declaracion' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Ano
                            </a>
                        </th>
                        <th data-column-index="7">
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="7" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=aum_pat_subito&sort_direction={% if current_order == 'aum_pat_subito' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Aum. Pat. Subito
                            </a>
                        </th>
                        
                        <th data-column-index="8"> {# Adjusted index #}
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="8" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=endeudamiento&sort_direction={% if current_order == 'endeudamiento' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                % Endeudamiento
                            </a>
                        </th>
                       
                        <th data-column-index="9"> {# Adjusted index #}
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="9" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=patrimonio&sort_direction={% if current_order == 'patrimonio' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Patrimonio
                            </a>
                        </th>
                        <th data-column-index="10"> {# Adjusted index #}
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="10" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=patrimonio_var_rel&sort_direction={% if current_order == 'patrimonio_var_rel' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Patrimonio Var. Rel. % 
                            </a>
                        </th>
                        <th data-column-index="11"> {# Adjusted index #}
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="11" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=patrimonio_var_abs&sort_direction={% if current_order == 'patrimonio_var_abs' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Patrimonio Var. Abs. $
                            </a>
                        </th>
                        <th data-column-index="12"> {# Adjusted index #}
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="12" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=activos&sort_direction={% if current_order == 'activos' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Activos
                            </a>
                        </th>
                        <th data-column-index="13"> {# Adjusted index #}
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="13" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=activos_var_rel&sort_direction={% if current_order == 'activos_var_rel' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Activos Var. Rel. %
                            </a>
                        </th>
                        <th data-column-index="14"> {# Adjusted index #}
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="14" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=activos_var_abs&sort_direction={% if current_order == 'activos_var_abs' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Activos Var. Abs. $
                            </a>
                        </th>
                        <th data-column-index="15"> {# Adjusted index #}
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="15" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=pasivos&sort_direction={% if current_order == 'pasivos' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Pasivos
                            </a>
                        </th>
                        <th data-column-index="16"> {# Adjusted index #}
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="16" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=pasivos_var_rel&sort_direction={% if current_order == 'pasivos_var_rel' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Pasivos Var. Rel. %
                            </a>
                        </th>
                        <th data-column-index="17"> {# Adjusted index #}
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="17" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=pasivos_var_abs&sort_direction={% if current_order == 'pasivos_var_abs' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Pasivos Var. Abs. $
                            </a>
                        </th>
                        <th data-column-index="18"> {# Adjusted index #}
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="18" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=cant_deudas&sort_direction={% if current_order == 'cant_deudas' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Cant. Deudas
                            </a>
                        </th>
                        <th data-column-index="19"> {# Adjusted index #}
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="19" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=ingresos&sort_direction={% if current_order == 'ingresos' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Ingresos
                            </a>
                        </th>
                        <th data-column-index="20"> {# Adjusted index #}
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="20" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=ingresos_var_rel&sort_direction={% if current_order == 'ingresos_var_rel' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Ingresos Var. Rel. %
                            </a>
                        </th>
                        <th data-column-index="21"> {# Adjusted index #}
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="21" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=ingresos_var_abs&sort_direction={% if current_order == 'ingresos_var_abs' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Ingresos Var. Abs. $
                            </a>
                        </th>
                        <th data-column-index="22"> {# Adjusted index #}
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="22" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=cant_ingresos&sort_direction={% if current_order == 'cant_ingresos' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Cant. Ingresos
                            </a>
                        </th>
                        <th data-column-index="23"> {# Adjusted index #}
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="23" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=banco_saldo&sort_direction={% if current_order == 'banco_saldo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Bancos Saldo
                            </a>
                        </th>
                        <th data-column-index="24"> {# Adjusted index #}
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="24" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=banco_saldo_var_rel&sort_direction={% if current_order == 'banco_saldo_var_rel' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Bancos Var. Rel. %
                            </a>
                        </th>
                        <th data-column-index="25"> {# Adjusted index #}
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="25" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=banco_saldo_var_abs&sort_direction={% if current_order == 'banco_saldo_var_abs' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Bancos Var. $
                            </a>
                        </th>
                        <th data-column-index="26"> {# Adjusted index #}
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="26" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=cant_cuentas&sort_direction={% if current_order == 'cant_cuentas' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Cant. Cuentas
                            </a>
                        </th>
                        <th data-column-index="27"> {# Adjusted index #}
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="27" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=cant_bancos&sort_direction={% if current_order == 'cant_bancos' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Cant. Bancos
                            </a>
                        </th>
                        <th data-column-index="28"> {# Adjusted index #}
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="28" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=bienes&sort_direction={% if current_order == 'bienes' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Bienes Valor
                            </a>
                        </th>
                        <th data-column-index="29"> {# Adjusted index #}
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="29" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=bienes_var_rel&sort_direction={% if current_order == 'bienes_var_rel' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Bienes Var. Rel. %
                            </a>
                        </th>
                        <th data-column-index="30"> {# Adjusted index #}
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="30" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=bienes_var_abs&sort_direction={% if current_order == 'bienes_var_abs' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Bienes Var. $
                            </a>
                        </th>
                        <th data-column-index="31"> {# Adjusted index #}
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="31" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=cant_bienes&sort_direction={% if current_order == 'cant_bienes' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Cant. Bienes
                            </a>
                        </th>
                        <th data-column-index="32"> {# Adjusted index #}
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="32" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=inversiones&sort_direction={% if current_order == 'inversiones' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Inversiones Valor
                            </a>
                        </th>
                        <th data-column-index="33"> {# Adjusted index #}
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="33" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=inversiones_var_rel&sort_direction={% if current_order == 'inversiones_var_rel' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Inversiones Var. Rel. %
                            </a>
                        </th>
                        <th data-column-index="34"> {# Adjusted index #}
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="34" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=inversiones_var_abs&sort_direction={% if current_order == 'inversiones_var_abs' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Inversiones Var. $
                            </a>
                        </th>
                        <th data-column-index="35"> {# Adjusted index #}
                            <button class="btn btn-sm btn-outline-secondary freeze-column-btn" data-column-index="35" title="Congelar columna">
                                <i class="fas fa-thumbtack"></i>
                            </button>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=cant_inversiones&sort_direction={% if current_order == 'cant_inversiones' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Cant. Inversiones
                            </a>
                        </th>
                        <th class="table-fixed-column" style="color: rgb(0, 0, 0);" data-column-index="36">Ver</th> {# Adjusted index #}
                    </tr>
                </thead>
                <tbody>
                    {% for report in financial_reports %}
                        <tr {% if report.person.revisar %}class="table-warning"{% endif %} data-person-cedula="{{ report.person.cedula }}" data-ano-declaracion="{{ report.ano_declaracion }}">
                            <td>
                                {# Revert to a standard form submission #}
                                <form action="{% url 'toggle_revisar_status' report.person.cedula %}" method="post" style="display:inline;">
                                    {% csrf_token %}
                                    <button type="submit" {# CHANGED BACK to type="submit" #}
                                            class="btn btn-link p-0 border-0 bg-transparent" {# Keep styling for icon appearance #}
                                            title="{% if report.person.revisar %}Desmarcar para Revisar{% else %}Marcar para Revisar{% endif %}">
                                        <i class="fas fa-{% if report.person.revisar %}check-square text-warning{% else %}square text-secondary{% endif %}"
                                        style="padding-left: 20px; font-size: 1.25rem;"></i>
                                    </button>
                                </form>
                            </td>

                            <td>{{ report.person.nombre_completo }}</td>
                            <td>{{ report.person.compania }}</td>
                            <td>{{ report.person.cargo }}</td>
                            <td>{{ report.person.comments|truncatechars:30|default:"" }}</td>
                            <td>{{ report.fk_id_periodo|floatformat:"0"|default:"-" }}</td>
                            <td>{{ report.ano_declaracion|floatformat:"0"|default:"-" }}</td>
                            
                            <td data-field="aum_pat_subito"
                                {% if report.aum_pat_subito >= 2 %}
                                    style="color: red;"
                                {% elif report.aum_pat_subito >= 1.5 and report.aum_pat_subito < 2 %}
                                    style="color: green;"
                                {% endif %}>
                                {{ report.aum_pat_subito|default:"0" }}
                            </td>
                            
                            <td data-field="endeudamiento"
                                {% if report.endeudamiento and report.endeudamiento > 70 %}
                                    style="color: red;"
                                {% elif report.endeudamiento and report.endeudamiento >= 50 and report.endeudamiento <= 70 %}
                                    style="color: green;"
                                {% endif %}>
                                {{ report.endeudamiento|floatformat:"2"|default:"0" }}
                            </td>

                            <td data-field="patrimonio">{{ report.patrimonio|floatformat:"0"|intcomma|default:"0" }}</td>
                            <td data-field="patrimonio_var_rel"
                                {% if report.patrimonio_var_rel and report.patrimonio_var_rel|floatformat:"0"|add:0 >= 50 or report.patrimonio_var_rel and report.patrimonio_var_rel|floatformat:"0"|add:0 <= -50 %}
                                    style="color: red;"
                                {% elif report.patrimonio_var_rel and report.patrimonio_var_rel|floatformat:"0"|add:0 >= 30 and report.patrimonio_var_rel|floatformat:"0"|add:0 < 50 %}
                                    style="color: green;"
                                {% elif report.patrimonio_var_rel and report.patrimonio_var_rel|floatformat:"0"|add:0 > -50 and report.patrimonio_var_rel|floatformat:"0"|add:0 <= -30 %}
                                    style="color: green;"
                                {% endif %}>
                                {{ report.patrimonio_var_rel|default:"0" }}
                            </td>

                            <td data-field="patrimonio_var_abs">{{ report.patrimonio_var_abs|floatformat:"0"|intcomma|default:"0" }}</td>
                            <td data-field="activos">{{ report.activos|floatformat:"0"|intcomma|default:"0" }}</td>
                            <td data-field="activos_var_rel"
                                {% if report.activos_var_rel and report.activos_var_rel|floatformat:"0"|add:0 >= 50 or report.activos_var_rel and report.activos_var_rel|floatformat:"0"|add:0 <= -50 %}
                                    style="color: red;"
                                {% elif report.activos_var_rel and report.activos_var_rel|floatformat:"0"|add:0 >= 30 and report.activos_var_rel|floatformat:"0"|add:0 < 50 %}
                                    style="color: green;"
                                {% elif report.activos_var_rel and report.activos_var_rel|floatformat:"0"|add:0 > -50 and report.activos_var_rel|floatformat:"0"|add:0 <= -30 %}
                                    style="color: green;"
                                {% endif %}>
                                {{ report.activos_var_rel|default:"0" }}
                            </td>
                            <td data-field="activos_var_abs">{{ report.activos_var_abs|floatformat:"0"|intcomma|default:"0" }}</td>
                            <td data-field="pasivos">{{ report.pasivos|floatformat:"0"|intcomma|default:"0" }}</td>
                            <td data-field="pasivos_var_rel"
                                {% if report.pasivos_var_rel and report.pasivos_var_rel|floatformat:"0"|add:0 >= 50 or report.pasivos_var_rel and report.pasivos_var_rel|floatformat:"0"|add:0 <= -50 %}
                                    style="color: red;"
                                {% elif report.pasivos_var_rel and report.pasivos_var_rel|floatformat:"0"|add:0 >= 30 and report.pasivos_var_rel|floatformat:"0"|add:0 < 50 %}
                                    style="color: green;"
                                {% elif report.pasivos_var_rel and report.pasivos_var_rel|floatformat:"0"|add:0 > -50 and report.pasivos_var_rel|floatformat:"0"|add:0 <= -30 %}
                                    style="color: green;"
                                {% endif %}>
                                {{ report.pasivos_var_rel|default:"0" }}
                            </td>
                            <td data-field="pasivos_var_abs">{{ report.pasivos_var_abs|floatformat:"0"|intcomma|default:"0" }}</td>
                            
                            <td data-field="cant_deudas"
                                {% if report.cant_deudas >= 6 %}
                                    style="color: red;"
                                {% elif report.cant_deudas >= 4 and report.cant_deudas < 6 %}
                                    style="color: green;"
                                {% endif %}>
                                {{ report.cant_deudas|floatformat:"0"|intcomma|default:"0" }}
                            </td>

                            <td data-field="ingresos">{{ report.ingresos|floatformat:"0"|intcomma|default:"0" }}</td>
                            <td data-field="ingresos_var_rel"
                                {% if report.ingresos_var_rel and report.ingresos_var_rel|floatformat:"0"|add:0 >= 50 or report.ingresos_var_rel and report.ingresos_var_rel|floatformat:"0"|add:0 <= -50 %}
                                    style="color: red;"
                                {% elif report.ingresos_var_rel and report.ingresos_var_rel|floatformat:"0"|add:0 >= 30 and report.ingresos_var_rel|floatformat:"0"|add:0 < 50 %}
                                    style="color: green;"
                                {% elif report.ingresos_var_rel and report.ingresos_var_rel|floatformat:"0"|add:0 > -50 and report.ingresos_var_rel|floatformat:"0"|add:0 <= -30 %}
                                    style="color: green;"
                                {% endif %}>
                                {{ report.ingresos_var_rel|default:"0" }}
                            </td>
                            <td data-field="ingresos_var_abs">{{ report.ingresos_var_abs|floatformat:"0"|intcomma|default:"0" }}</td>
                            <td data-field="cant_ingresos">{{ report.cant_ingresos|floatformat:"0"|intcomma|default:"0" }}</td>
                            <td data-field="banco_saldo">{{ report.banco_saldo|floatformat:"0"|intcomma|default:"0" }}</td>
                            <td data-field="banco_saldo_var_rel"
                                {% if report.banco_saldo_var_rel and report.banco_saldo_var_rel|floatformat:"0"|add:0 >= 50 or report.banco_saldo_var_rel and report.banco_saldo_var_rel|floatformat:"0"|add:0 <= -50 %}
                                    style="color: red;"
                                {% elif report.banco_saldo_var_rel and report.banco_saldo_var_rel|floatformat:"0"|add:0 >= 30 and report.banco_saldo_var_rel|floatformat:"0"|add:0 < 50 %}
                                    style="color: green;"
                                {% elif report.banco_saldo_var_rel and report.banco_saldo_var_rel|floatformat:"0"|add:0 > -50 and report.banco_saldo_var_rel|floatformat:"0"|add:0 <= -30 %}
                                    style="color: green;"
                                {% endif %}>
                                {{ report.banco_saldo_var_rel|default:"0" }}
                            </td>
                            <td data-field="banco_saldo_var_abs">{{ report.banco_saldo_var_abs|floatformat:"0"|intcomma|default:"0" }}</td>
                            
                            <td data-field="cant_cuentas"
                                {% if report.cant_cuentas >= 6 %}
                                    style="color: red;"
                                {% elif report.cant_cuentas >= 4 and report.cant_cuentas < 6 %}
                                    style="color: green;"
                                {% endif %}>
                                {{ report.cant_cuentas|floatformat:"0"|intcomma|default:"0" }}
                            </td>
                            
                            <td data-field="cant_bancos">{{ report.cant_bancos|floatformat:"0"|intcomma|default:"0" }}</td>
                            <td data-field="bienes">{{ report.bienes|floatformat:"0"|intcomma|default:"0" }}</td>
                            <td data-field="bienes_var_rel"
                                {% if report.bienes_var_rel and report.bienes_var_rel|floatformat:"0"|add:0 >= 50 or report.bienes_var_rel and report.bienes_var_rel|floatformat:"0"|add:0 <= -50 %}
                                    style="color: red;"
                                {% elif report.bienes_var_rel and report.bienes_var_rel|floatformat:"0"|add:0 >= 30 and report.bienes_var_rel|floatformat:"0"|add:0 < 50 %}
                                    style="color: green;"
                                {% elif report.bienes_var_rel and report.bienes_var_rel|floatformat:"0"|add:0 > -50 and report.bienes_var_rel|floatformat:"0"|add:0 <= -30 %}
                                    style="color: green;"
                                {% endif %}>
                                {{ report.bienes_var_rel|default:"0" }}
                            </td>
                            <td data-field="bienes_var_abs">{{ report.bienes_var_abs|floatformat:"0"|intcomma|default:"0" }}</td>
                            <td data-field="cant_bienes">{{ report.cant_bienes|floatformat:"0"|intcomma|default:"0" }}</td>
                            <td data-field="inversiones">{{ report.inversiones|floatformat:"0"|intcomma|default:"0" }}</td>
                            <td data-field="inversiones_var_rel"
                                {% if report.inversiones_var_rel and report.inversiones_var_rel|floatformat:"0"|add:0 >= 50 or report.inversiones_var_rel and report.inversiones_var_rel|floatformat:"0"|add:0 <= -50 %}
                                    style="color: red;"
                                {% elif report.inversiones_var_rel and report.inversiones_var_rel|floatformat:"0"|add:0 >= 30 and report.inversiones_var_rel|floatformat:"0"|add:0 < 50 %}
                                    style="color: green;"
                                {% elif report.inversiones_var_rel and report.inversiones_var_rel|floatformat:"0"|add:0 > -50 and report.inversiones_var_rel|floatformat:"0"|add:0 <= -30 %}
                                    style="color: green;"
                                {% endif %}>
                                {{ report.inversiones_var_rel|default:"0" }}
                            </td>
                            <td data-field="inversiones_var_abs">{{ report.inversiones_var_abs|floatformat:"0"|intcomma|default:"0" }}</td>
                            <td data-field="cant_inversiones">{{ report.cant_inversiones|floatformat:"0"|intcomma|default:"0" }}</td>
                            
                            <td class="table-fixed-column">
                                <a href="{% url 'person_details' report.person.cedula %}"
                                class="btn btn-custom-primary btn-sm"
                                title="View person details">
                                    <i class="bi bi-person-vcard-fill"></i>
                                </a>
                            </td>
                        </tr>
                    {% empty %}
                        <tr>
                            <td colspan="37" class="text-center py-4"> {# Adjusted colspan #}
                                {% if request.GET.q or request.GET.column_0 or request.GET.operator_0 or request.GET.value_0 %}
                                    Sin reportes financieros que coincidan con los filtros.
                                {% else %}
                                    Sin reportes financieros
                                {% endif %}
                            </td>
                        </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
        
        {% if page_obj.has_other_pages %}
        <div class="p-3">
            <nav aria-label="Page navigation">
                <ul class="pagination justify-content-center">
                    {% if page_obj.has_previous %}
                        <li class="page-item">
                            <a class="page-link" href="?page=1{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="First">
                                <span aria-hidden="true">&laquo;&laquo;</span>
                            </a>
                        </li>
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.previous_page_number }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Previous">
                                <span aria-hidden="true">&laquo;</span>
                            </a>
                        </li>
                    {% endif %}
                    
                    {% for num in page_obj.paginator.page_range %}
                        {% if page_obj.number == num %}
                            <li class="page-item active"><a class="page-link" href="#">{{ num }}</a></li>
                        {% elif num > page_obj.number|add:'-3' and num < page_obj.number|add:'3' %}
                            <li class="page-item"><a class="page-link" href="?page={{ num }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}">{{ num }}</a></li>
                        {% endif %}
                    {% endfor %}
                    
                    {% if page_obj.has_next %}
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.next_page_number }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Next">
                                <span aria-hidden="true">&raquo;</span>
                            </a>
                        </li>
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.paginator.num_pages }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Last">
                                <span aria-hidden="true">&raquo;&raquo;</span>
                            </a>
                        </li>
                    {% endif %}
                </ul>
            </nav>
        </div>
        {% endif %}
    </div>

<script>
    let filterCount = 0; // Global counter for filter groups

    // Function to show/hide value2 input based on operator
    function toggleValueInput(selectElement) {
        const filterGroup = selectElement.closest('.filter-group');
        const value1Container = filterGroup.querySelector('.value1-container');
        const value1Input = filterGroup.querySelector('.value1-input');
        const value2Container = filterGroup.querySelector('.value2-container');
        const operatorSelect = selectElement;

        if (operatorSelect.value === 'between') {
            value2Container.style.display = 'block';
            value1Input.placeholder = "Primer Valor";
            value1Container.classList.remove('col-md-3');
            value1Container.classList.add('col-md-2');
        } else {
            value2Container.style.display = 'none';
            value2Container.querySelector('.value2-input').value = ''; // Clear value2 when hidden
            value1Input.placeholder = "Valor";
            value1Container.classList.remove('col-md-2');
            value1Container.classList.add('col-md-3');
        }
    }

    // Function to add a new filter group
    function addFilterGroup(initialValues = {}) {
        const filterContainer = document.getElementById('filter-rows-container');
        const template = document.querySelector('.filter-group-template .filter-group');
        const newFilterGroup = template.cloneNode(true);
        newFilterGroup.style.display = 'flex'; // Make the cloned element visible

        const currentIndex = filterCount++;
        
        newFilterGroup.querySelectorAll('select, input').forEach(input => {
            const oldName = input.name;
            // Replace '_X' placeholder with actual index
            input.name = oldName.replace('_X', '_' + currentIndex);
            input.id = input.name; // Assign an ID based on the name

            // Set initial values if provided
            if (initialValues[input.name]) {
                input.value = initialValues[input.name];
            } else {
                input.value = ''; // Clear values for newly added empty filters
            }
            
            // Reset placeholder for value1 inputs
            if (input.classList.contains('value1-input')) {
                input.placeholder = "Valor";
            }
        });

        // Set up event listeners for the new filter group
        const newOperatorSelect = newFilterGroup.querySelector('.operator-select');
        newOperatorSelect.onchange = function() { toggleValueInput(this); };

        const removeButton = newFilterGroup.querySelector('.remove-filter-btn');
        removeButton.onclick = function() {
            newFilterGroup.remove();
        };

        filterContainer.appendChild(newFilterGroup);

        // Apply toggleValueInput logic based on the operator's initial value (if loaded from GET)
        toggleValueInput(newOperatorSelect);
    }

    document.addEventListener('DOMContentLoaded', function() {
        const urlParams = new URLSearchParams(window.location.search);
        let hasFilters = false;
        
        // This loop checks for the existence of 'column_i' or 'value_i' to determine if a filter was applied.
        let tempFilterIndex = 0;
        while (urlParams.has('column_' + tempFilterIndex) || 
               urlParams.has('operator_' + tempFilterIndex) || 
               urlParams.has('value_' + tempFilterIndex) ||
               urlParams.has('value2_' + tempFilterIndex)) {
            
            const initialValues = {};
            initialValues['column_' + tempFilterIndex] = urlParams.get('column_' + tempFilterIndex);
            initialValues['operator_' + tempFilterIndex] = urlParams.get('operator_' + tempFilterIndex);
            initialValues['value_' + tempFilterIndex] = urlParams.get('value_' + tempFilterIndex);
            initialValues['value2_' + tempFilterIndex] = urlParams.get('value2_' + tempFilterIndex);

            addFilterGroup(initialValues);
            hasFilters = true;
            tempFilterIndex++;
        }

        // If no filters were present in the URL, add one blank filter group
        if (!hasFilters) {
            addFilterGroup();
        }

        // Add event listener for the 'Add Filter' button
        const addFilterButton = document.getElementById('add-filter-btn');
        if (addFilterButton) {
            addFilterButton.addEventListener('click', function() {
                addFilterGroup({});
            });
        }
    });
</script>

</div>
{% endblock %}
"@ | Out-File -FilePath "core/templates/finances.html" -Encoding utf8

# details template
@" 
{% extends "master.html" %}
{% load static %}
{% load humanize %}

{% block title %}Detalles - {{ myperson.nombre_completo }}{% endblock %}
{% block navbar_title %}{{ myperson.nombre_completo }}{% endblock %}

{% block navbar_buttons %}
<div>
    <a href="/admin/core/person/{{ myperson.cedula }}/change/" class="btn btn-outline-dark" title="Admin">
        <i class="fas fa-pencil-alt"></i>
    </a>
    <a href="/" class="btn btn-custom-primary">
        <i class="fas fa-chart-pie" style="color: rgb(255, 111, 0);"></i>
    </a>
    <a href="{% url 'person_list' %}" class="btn btn-custom-primary">
        <i class="fas fa-users"></i>
    </a>
    <a href="{% url 'financial_report_list' %}" class="btn btn-custom-primary" title="Bienes y Rentas">
        <i class="fas fa-chart-line" style="color: green;"></i>
    </a>
    <a href="{% url 'tcs_list' %}" class="btn btn-custom-primary" title="Tarjetas">
        <i class="far fa-credit-card" style="color: blue;"></i>
    </a>
    <a href="{% url 'conflict_list' %}" class="btn btn-custom-primary">
        <i class="fas fa-balance-scale" style="color: orange;"></i>
    </a>
    <a href="{% url 'alerts_list' %}" class="btn btn-custom-primary" title="Alertas">
        {% if alerts_count > 0 %}
            <span class="badge bg-danger">{{ alerts_count }}</span>
        {% else %}
            <span class="badge bg-secondary">0</span>
        {% endif %}
        <i class="fas fa-bell" style="color: red;"></i>
    </a>
    <a href="{% url 'import' %}" class="btn btn-custom-primary">
        <i class="fas fa-database"></i> 
    </a>
    <form method="post" action="{% url 'logout' %}" class="d-inline">
        {% csrf_token %}
        <button type="submit" class="btn btn-custom-primary" title="Cerrar sesion">
            <i class="fas fa-sign-out-alt"></i>
        </button>
    </form>
</div>
{% endblock %}

{% block content %}
<div class="row">
    <div class="col-md-6 mb-4"> {# Column for Informacion Personal - half width #}
        <div class="card h-100"> {# Added h-100 for equal height #}
            <div class="card-header bg-light">
                <h5 class="mb-0">Informacion Personal</h5>
            </div>
            <div class="card-body">
                <table class="table">
                    <tr>
                        <th>ID:</th>
                        <td>{{ myperson.cedula }}</td>
                    </tr>
                    <tr>
                        <th>Nombre:</th>
                        <td>{{ myperson.nombre_completo }}</td>
                    </tr>
                    <tr>
                        <th>Cargo:</th>
                        <td>{{ myperson.cargo }}</td>
                    </tr>
                    <tr>
                        <th>Correo:</th>
                        <td>{{ myperson.correo }}</td>
                    </tr>
                    <tr>
                        <th>Compania:</th>
                        <td>{{ myperson.compania }}</td>
                    </tr>
                    <tr>
                        <th>Estado:</th>
                        <td>
                            <span class="badge bg-{% if myperson.estado == 'Activo' %}success{% else %}danger{% endif %}">
                                {{ myperson.estado }}
                            </span>
                        </td>
                    </tr>
                    <tr>
                        <th>Por revisar:</th>
                        <td>
                            {% if myperson.revisar %}
                                <span class="badge bg-warning text-dark">Si</span>
                            {% else %}
                                <span class="badge bg-secondary">No</span>
                            {% endif %}
                        </td>
                    </tr>
                    <tr>
                        <th>Comentarios:</th>
                        <td>{{ myperson.comments|linebreaks }}</td>
                    </tr>
                </table>
            </div>
        </div>
    </div>

    <div class="col-md-6 mb-4"> {# Column for Conflictos Declarados - half width #}
        <div class="card h-100"> {# Added h-100 for equal height #}
            <div class="card-header bg-light d-flex justify-content-between align-items-center">
                <h5 class="mb-0">Conflictos Declarados</h5>
                {% if conflicts %}
                <div class="form-group">
                    <label for="conflict-select" class="sr-only">Seleccionar Conflicto</label>
                    <select class="form-control" id="conflict-select">
                        {% for conflict in conflicts %}
                            <option value="{{ forloop.counter0 }}">Fecha Declaracion: {{ conflict.fecha_inicio|date:"d/m/Y"|default:"-" }}</option>
                        {% endfor %}
                    </select>
                </div>
                {% endif %}
            </div>
            <div class="card-body p-0">
                {% if conflicts %}
                {% for conflict in conflicts %}
                <div class="conflict-table" id="conflict-table-{{ forloop.counter0 }}" {% if not forloop.first %}style="display: none;"{% endif %}>
                    <div class="table-responsive">
                        <table class="table table-striped table-hover mb-0">
                            <tbody>
                                <tr>
                                    <th scope="row">Fecha Declaracion</th>
                                    <td>{{ conflict.fecha_inicio|date:"d/m/Y"|default:"-" }}</td>
                                </tr>
                                <tr>
                                    <th scope="row">Accionista de algun proveedor del grupo</th>
                                    <td class="text-center">{% if conflict.q1 %}<i style="color: red;">SI</i>{% else %}<i style="color: green;">NO</i>{% endif %}</td>
                                </tr>
                                {% if conflict.q1_detalle and conflict.q1_detalle|lower != "nan" %}
                                <tr>
                                    <th></th> {# Empty header for alignment #}
                                    <td><small class="text-muted">{{ conflict.q1_detalle|linebreaksbr }}</small></td>
                                </tr>
                                {% endif %}
                                <tr>
                                    <th scope="row">Familiar de algun accionista, proveedor o empleado</th>
                                    <td class="text-center">{% if conflict.q2 %}<i style="color: red;">SI</i>{% else %}<i style="color: green;">NO</i>{% endif %}</td>
                                </tr>
                                {% if conflict.q2_detalle and conflict.q2_detalle|lower != "nan" %}
                                <tr>
                                    <th></th>
                                    <td><small class="text-muted">{{ conflict.q2_detalle|linebreaksbr }}</small></td>
                                </tr>
                                {% endif %}
                                <tr>
                                    <th scope="row">Accionista de alguna compania del grupo</th>
                                    <td class="text-center">{% if conflict.q3 %}<i style="color: red;">SI</i>{% else %}<i style="color: green;">NO</i>{% endif %}</td>
                                </tr>
                                {% if conflict.q3_detalle and conflict.q3_detalle|lower != "nan" %}
                                <tr>
                                    <th></th>
                                    <td><small class="text-muted">{{ conflict.q3_detalle|linebreaksbr }}</small></td>
                                </tr>
                                {% endif %}
                                <tr>
                                    <th scope="row">Actividades extralaborales</th>
                                    <td class="text-center">{% if conflict.q4 %}<i style="color: red;">SI</i>{% else %}<i style="color: green;">NO</i>{% endif %}</td>
                                </tr>
                                {% if conflict.q4_detalle and conflict.q4_detalle|lower != "nan" %}
                                <tr>
                                    <th></th>
                                    <td><small class="text-muted">{{ conflict.q4_detalle|linebreaksbr }}</small></td>
                                </tr>
                                {% endif %}
                                <tr>
                                    <th scope="row">Negocios o bienes con empleados del grupo</th>
                                    <td class="text-center">{% if conflict.q5 %}<i style="color: red;">SI</i>{% else %}<i style="color: green;">NO</i>{% endif %}</td>
                                </tr>
                                {% if conflict.q5_detalle and conflict.q5_detalle|lower != "nan" %}
                                <tr>
                                    <th></th>
                                    <td><small class="text-muted">{{ conflict.q5_detalle|linebreaksbr }}</small></td>
                                </tr>
                                {% endif %}
                                <tr>
                                    <th scope="row">Participacion en juntas o consejos directivos</th>
                                    <td class="text-center">{% if conflict.q6 %}<i style="color: red;">SI</i>{% else %}<i style="color: green;">NO</i>{% endif %}</td>
                                </tr>
                                {% if conflict.q6_detalle and conflict.q6_detalle|lower != "nan" %}
                                <tr>
                                    <th></th>
                                    <td><small class="text-muted">{{ conflict.q6_detalle|linebreaksbr }}</small></td>
                                </tr>
                                {% endif %}
                                <tr>
                                    <th scope="row">Potencial conflicto diferente a los anteriores</th>
                                    <td class="text-center">{% if conflict.q7 %}<i style="color: red;">SI</i>{% else %}<i style="color: green;">NO</i>{% endif %}</td>
                                </tr>
                                {% if conflict.q7_detalle and conflict.q7_detalle|lower != "nan" %}
                                <tr>
                                    <th></th>
                                    <td><small class="text-muted">{{ conflict.q7_detalle|linebreaksbr }}</small></td>
                                </tr>
                                {% endif %}
                                <tr>
                                    <th scope="row">Consciente del codigo de conducta empresarial</th>
                                    <td class="text-center">{% if conflict.q8 %}<i style="color: green;">SI</i>{% else %}<i style="color: red;">NO</i>{% endif %}</td>
                                </tr>
                                {# Q8 and Q9 do not have detail fields, so no new rows for them #}
                                <tr>
                                    <th scope="row">Veracidad de la informacion consignada</th>
                                    <td class="text-center">{% if conflict.q9 %}<i style="color: green;">SI</i>{% else %}<i style="color: RED;">NO</i>{% endif %}</td>
                                </tr>
                                <tr>
                                    <th scope="row">Familiar de algun funcionario publico</th>
                                    <td class="text-center">{% if conflict.q10 %}<i style="color: red;">SI</i>{% else %}<i style="color: green;">NO</i>{% endif %}</td>
                                </tr>
                                {% if conflict.q10_detalle and conflict.q10_detalle|lower != "nan" %}
                                <tr>
                                    <th></th>
                                    <td><small class="text-muted">{{ conflict.q10_detalle|linebreaksbr }}</small></td>
                                </tr>
                                {% endif %}
                                <tr>
                                    <th scope="row">Relacion con el sector publico o funcionario publico</th>
                                    <td class="text-center">{% if conflict.q11 %}<i style="color: red;">SI</i>{% else %}<i style="color: green;">NO</i>{% endif %}</td>
                                </tr>
                                {% if conflict.q11_detalle and conflict.q11_detalle|lower != "nan" %}
                                <tr>
                                    <th></th>
                                    <td><small class="text-muted">{{ conflict.q11_detalle|linebreaksbr }}</small></td>
                                </tr>
                                {% endif %}
                            </tbody>
                        </table>
                    </div>
                </div>
                {% endfor %}
                {% else %}
                <p class="text-center py-4">No hay conflictos declarados disponibles</p>
                {% endif %}
            </div>
        </div>
    </div>
</div>

<div class="row"> {# New row for Reportes Financieros #}
    <div class="col-md-12 mb-4"> {# Full width column for Reportes Financieros #}
        <div class="card">
            <div class="card-header bg-light d-flex justify-content-between align-items-center">
                <h5 class="mb-0">Reportes Financieros</h5>
                <div>
                    <span class="badge bg-primary">{{ financial_reports.count }} periodos</span>
                </div>
            </div>
            <div class="card-body p-0">
                <div class="table-responsive table-container">
                    <table class="table table-striped table-hover mb-0">
                        <thead class="table-fixed-header">
                            <tr>
                                <th>Ano</th>
                                <th scope="col">Variaciones</th>
                                <th>Activos</th>
                                <th>Pasivos</th>
                                <th>Ingresos</th>
                                <th>Patrimonio</th>
                                <th>Banco</th>
                                <th>Bienes</th>
                                <th>Inversiones</th>
                                <th>% Apalancamiento</th>
                                <th>% Endeudamiento</th>
                                <th>Indice</th>
                            </tr>
                        </thead>
                        <tbody>
                            {% for report in financial_reports %}
                            <tr data-person-cedula="{{ myperson.cedula }}" data-ano-declaracion="{{ report.ano_declaracion|floatformat:"0" }}">
                                <td>{{ report.ano_declaracion|floatformat:"0"|default:"-" }}</td>
                                <th>Relativa</th>
                                <td data-field="activos_var_rel"
                                    {% if report.activos_var_rel and report.activos_var_rel|floatformat:"0"|add:0 >= 50 or report.activos_var_rel and report.activos_var_rel|floatformat:"0"|add:0 <= -50 %}
                                        style="color: red;"
                                    {% elif report.activos_var_rel and report.activos_var_rel|floatformat:"0"|add:0 >= 30 and report.activos_var_rel|floatformat:"0"|add:0 < 50 %}
                                        style="color: green;"
                                    {% elif report.activos_var_rel and report.activos_var_rel|floatformat:"0"|add:0 > -50 and report.activos_var_rel|floatformat:"0"|add:0 <= -30 %}
                                        style="color: green;"
                                    {% endif %}>
                                    {{ report.activos_var_rel|default:"0" }}
                                </td>
                                <td data-field="pasivos_var_rel"
                                    {% if report.pasivos_var_rel and report.pasivos_var_rel|floatformat:"0"|add:0 >= 50 or report.pasivos_var_rel and report.pasivos_var_rel|floatformat:"0"|add:0 <= -50 %}
                                        style="color: red;"
                                    {% elif report.pasivos_var_rel and report.pasivos_var_rel|floatformat:"0"|add:0 >= 30 and report.pasivos_var_rel|floatformat:"0"|add:0 < 50 %}
                                        style="color: green;"
                                    {% elif report.pasivos_var_rel and report.pasivos_var_rel|floatformat:"0"|add:0 > -50 and report.pasivos_var_rel|floatformat:"0"|add:0 <= -30 %}
                                        style="color: green;"
                                    {% endif %}>
                                    {{ report.pasivos_var_rel|default:"0" }}
                                </td>
                                <td data-field="ingresos_var_rel"
                                    {% if report.ingresos_var_rel and report.ingresos_var_rel|floatformat:"0"|add:0 >= 50 or report.ingresos_var_rel and report.ingresos_var_rel|floatformat:"0"|add:0 <= -50 %}
                                        style="color: red;"
                                    {% elif report.ingresos_var_rel and report.ingresos_var_rel|floatformat:"0"|add:0 >= 30 and report.ingresos_var_rel|floatformat:"0"|add:0 < 50 %}
                                        style="color: green;"
                                    {% elif report.ingresos_var_rel and report.ingresos_var_rel|floatformat:"0"|add:0 > -50 and report.ingresos_var_rel|floatformat:"0"|add:0 <= -30 %}
                                        style="color: green;"
                                    {% endif %}>
                                    {{ report.ingresos_var_rel|default:"0" }}
                                </td>
                                <td data-field="patrimonio_var_rel"
                                    {% if report.patrimonio_var_rel and report.patrimonio_var_rel|floatformat:"0"|add:0 >= 50 or report.patrimonio_var_rel and report.patrimonio_var_rel|floatformat:"0"|add:0 <= -50 %}
                                        style="color: red;"
                                    {% elif report.patrimonio_var_rel and report.patrimonio_var_rel|floatformat:"0"|add:0 >= 30 and report.patrimonio_var_rel|floatformat:"0"|add:0 < 50 %}
                                        style="color: green;"
                                    {% elif report.patrimonio_var_rel and report.patrimonio_var_rel|floatformat:"0"|add:0 > -50 and report.patrimonio_var_rel|floatformat:"0"|add:0 <= -30 %}
                                        style="color: green;"
                                    {% endif %}>
                                    {{ report.patrimonio_var_rel|default:"0" }}
                                </td>
                                <td data-field="banco_saldo_var_rel"
                                    {% if report.banco_saldo_var_rel and report.banco_saldo_var_rel|floatformat:"0"|add:0 >= 50 or report.banco_saldo_var_rel and report.banco_saldo_var_rel|floatformat:"0"|add:0 <= -50 %}
                                        style="color: red;"
                                    {% elif report.banco_saldo_var_rel and report.banco_saldo_var_rel|floatformat:"0"|add:0 >= 30 and report.banco_saldo_var_rel|floatformat:"0"|add:0 < 50 %}
                                        style="color: green;"
                                    {% elif report.banco_saldo_var_rel and report.banco_saldo_var_rel|floatformat:"0"|add:0 > -50 and report.banco_saldo_var_rel|floatformat:"0"|add:0 <= -30 %}
                                        style="color: green;"
                                    {% endif %}>
                                    {{ report.banco_saldo_var_rel|default:"0" }}
                                </td>
                                <td data-field="bienes_var_rel"
                                    {% if report.bienes_var_rel and report.bienes_var_rel|floatformat:"0"|add:0 >= 50 or report.bienes_var_rel and report.bienes_var_rel|floatformat:"0"|add:0 <= -50 %}
                                        style="color: red;"
                                    {% elif report.bienes_var_rel and report.bienes_var_rel|floatformat:"0"|add:0 >= 30 and report.bienes_var_rel|floatformat:"0"|add:0 < 50 %}
                                        style="color: green;"
                                    {% elif report.bienes_var_rel and report.bienes_var_rel|floatformat:"0"|add:0 > -50 and report.bienes_var_rel|floatformat:"0"|add:0 <= -30 %}
                                        style="color: green;"
                                    {% endif %}>
                                    {{ report.bienes_var_rel|default:"0" }}
                                </td>
                                <td data-field="inversiones_var_rel"
                                    {% if report.inversiones_var_rel and report.inversiones_var_rel|floatformat:"0"|add:0 >= 50 or report.inversiones_var_rel and report.inversiones_var_rel|floatformat:"0"|add:0 <= -50 %}
                                        style="color: red;"
                                    {% elif report.inversiones_var_rel and report.inversiones_var_rel|floatformat:"0"|add:0 >= 30 and report.inversiones_var_rel|floatformat:"0"|add:0 < 50 %}
                                        style="color: green;"
                                    {% elif report.inversiones_var_rel and report.inversiones_var_rel|floatformat:"0"|add:0 > -50 and report.inversiones_var_rel|floatformat:"0"|add:0 <= -30 %}
                                        style="color: green;"
                                    {% endif %}>
                                    {{ report.inversiones_var_rel|default:"0" }}
                                </td>
                                <td data-field="apalancamiento_var_rel">{{ report.apalancamiento_var_rel|default:"0" }}</td>
                                <td data-field="endeudamiento_var_rel">{{ report.endeudamiento_var_rel|default:"0" }}</td>
                                <td data-field="aum_pat_subito"
                                    {% if report.aum_pat_subito >= 2 %}
                                        style="color: red;"
                                    {% elif report.aum_pat_subito >= 1.5 and report.aum_pat_subito < 2 %}
                                        style="color: green;"
                                    {% endif %}>
                                    {{ report.aum_pat_subito|default:"0" }}
                                </td>
                            </tr>
                            <tr>
                                <th></th>
                                <th scope="col">Absoluta</th>
                                <td data-field="activos_var_abs">{{ report.activos_var_abs|floatformat:"0"|intcomma|default:"0" }}</td>
                                <td data-field="pasivos_var_abs">{{ report.pasivos_var_abs|floatformat:"0"|intcomma|default:"0" }}</td>
                                <td data-field="ingresos_var_abs">{{ report.ingresos_var_abs|floatformat:"0"|intcomma|default:"0" }}</td>
                                <td data-field="patrimonio_var_abs">{{ report.patrimonio_var_abs|floatformat:"0"|intcomma|default:"0" }}</td>
                                <td data-field="banco_saldo_var_abs">{{ report.banco_saldo_var_abs|floatformat:"0"|intcomma|default:"0" }}</td>
                                <td data-field="bienes_var_abs">{{ report.bienes_var_abs|floatformat:"0"|intcomma|default:"0" }}</td>
                                <td data-field="inversiones_var_abs">{{ report.inversiones_var_abs|floatformat:"0"|intcomma|default:"0" }}</td>
                                <td data-field="apalancamiento_var_abs">{{ report.apalancamiento_var_abs|default:"0" }}</td>
                                <td data-field="endeudamiento_var_abs">{{ report.endeudamiento_var_abs|default:"0" }}</td>
                                <td data-field="capital_var_abs">{{ report.capital_var_abs|floatformat:"0"|intcomma|default:"0" }}</td>
                            </tr>
                            <tr>
                                <td></td>
                                <th scope="col">Total</th>
                                <td data-field="activos">&#36;{{ report.activos|floatformat:"0"|intcomma|default:"0" }}</td>
                                <td data-field="pasivos">&#36;{{ report.pasivos|floatformat:"0"|intcomma|default:"0" }}</td>
                                <td data-field="ingresos">&#36;{{ report.ingresos|floatformat:"0"|intcomma|default:"0" }}</td>
                                <td data-field="patrimonio">&#36;{{ report.patrimonio|floatformat:"0"|intcomma|default:"0" }}</td>
                                <td data-field="banco_saldo">&#36;{{ report.banco_saldo|floatformat:"0"|intcomma|default:"0" }}</td>
                                <td data-field="bienes">&#36;{{ report.bienes|floatformat:"0"|intcomma|default:"0" }}</td>
                                <td data-field="inversiones">&#36;{{ report.inversiones|floatformat:"0"|intcomma|default:"0" }}</td>
                                <td data-field="apalancamiento">{{ report.apalancamiento|floatformat:2|default:"0" }}</td>

                                <td data-field="endeudamiento"
                                    {% if report.endeudamiento and report.endeudamiento > 70 %}
                                        style="color: red;"
                                    {% elif report.endeudamiento and report.endeudamiento >= 50 and report.endeudamiento <= 70 %}
                                        style="color: green;"
                                    {% endif %}>
                                    {{ report.endeudamiento|floatformat:"2"|default:"0" }}
                                </td>
                                
                                <td data-field="capital">&#36;{{ report.capital|floatformat:"0"|intcomma|default:"0" }}</td>
                            </tr>
                            <tr>
                                <th></th>
                                <th scope="col">Cant.</th>
                                <td>0</td>
                                <td data-field="cant_deudas"
                                    {% if report.cant_deudas >= 6 %}
                                        style="color: red;"
                                    {% elif report.cant_deudas >= 4 and report.cant_deudas < 6 %}
                                        style="color: green;"
                                    {% endif %}>
                                    {{ report.cant_deudas|floatformat:"0"|intcomma|default:"0" }}
                                </td>
                                <td data-field="cant_ingresos">{{ report.cant_ingresos|floatformat:"0"|intcomma|default:"0" }}</td>
                                <td data-field="cant_cuentas">{{ report.cant_cuentas|floatformat:"0"|intcomma|default:"0" }}</td>
                                <td data-field="cant_bancos">B{{ report.cant_bancos|floatformat:"0"|intcomma|default:"0" }} C{{ report.cant_cuentas|floatformat:"0"|intcomma|default:"0" }}</td>
                                <td data-field="cant_bienes">{{ report.cant_bienes|floatformat:"0"|intcomma|default:"0" }}</td>
                                <td data-field="cant_inversiones">{{ report.cant_inversiones|floatformat:"0"|intcomma|default:"0" }}</td>
                                <td></td>
                                <td></td>
                                <td></td>
                            </tr>
                            {% empty %}
                            <tr>
                                <td colspan="12" class="text-center py-4">
                                    No hay reportes financieros disponibles
                                </td>
                            </tr>
                            {% endfor %}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
</div>
<script src="{% static 'js/details.js' %}"></script>
<script>
    document.addEventListener('DOMContentLoaded', function() {
        const conflictSelect = document.getElementById('conflict-select');
        const conflictTables = document.querySelectorAll('.conflict-table');

        if (conflictSelect) {
            conflictSelect.addEventListener('change', function() {
                const selectedIndex = this.value;
                conflictTables.forEach((table, index) => {
                    if (index == selectedIndex) {
                        table.style.display = 'block';
                    } else {
                        table.style.display = 'none';
                    }
                });
            });
        }
    });
</script>
{% endblock %}
"@ | Out-File -FilePath "core/templates/details.html" -Encoding utf8

# Create alert template
@"
{% extends "master.html" %}
{% load static %}

{% block title %}Alertas{% endblock %}
{% block navbar_title %}Alertas{% endblock %}

{% block navbar_buttons %}
<div>
    <a href="/" class="btn btn-custom-primary">
        <i class="fas fa-chart-pie" style="color: rgb(255, 111, 0);"></i>
    </a>
    <a href="{% url 'financial_report_list' %}" class="btn btn-custom-primary" title="Bienes y Rentas">
        <i class="fas fa-chart-line" style="color: green;"></i>
    </a>
    <a href="{% url 'tcs_list' %}" class="btn btn-custom-primary" title="Tarjetas">
        <i class="far fa-credit-card" style="color: blue;"></i>
    </a>
    <a href="{% url 'conflict_list' %}" class="btn btn-custom-primary">
        <i class="fas fa-balance-scale" style="color: orange;"></i>
    </a>
    <a href="{% url 'alerts_list' %}" class="btn btn-custom-primary" title="Alertas">
        {% if alerts_count > 0 %}
            <span class="badge bg-danger">{{ alerts_count }}</span>
        {% else %}
            <span class="badge bg-secondary">0</span>
        {% endif %}
        <i class="fas fa-bell" style="color: red;"></i>
    </a>
    <a href="{% url 'import' %}" class="btn btn-custom-primary">
        <i class="fas fa-database"></i> 
    </a>
    <a href="{% url 'export_persons_excel' %}{% if request.GET %}?{{ request.GET.urlencode }}{% endif %}" class="btn btn-custom-primary">
        <i class="fas fa-file-excel" style="color: green;"></i>
    </a>
    <form method="post" action="{% url 'logout' %}" class="d-inline">
        {% csrf_token %}
        <button type="submit" class="btn btn-custom-primary" title="Cerrar sesion">
            <i class="fas fa-sign-out-alt"></i>
        </button>
    </form>
</div>
{% endblock %}

{% block content %}
<!-- Search Form (Optional, but good for consistency) -->
<div class="card mb-4 border-0 shadow" style="background-color:rgb(224, 224, 224);">
    <div class="card-body">
        <form method="get" action="." class="row g-3 align-items-center">
            <div class="d-flex align-items-center">
                <span class="badge bg-success">
                    {{ page_obj.paginator.count }} alertas
                </span>
            </div>
            <!-- General Search -->
            <div class="col-md-4">
                <input type="text" 
                       name="q" 
                       class="form-control form-control-lg" 
                       placeholder="Buscar persona o cedula" 
                       value="{{ request.GET.q }}">
            </div>
            
            <!-- Submit Buttons -->
            <div class="col-md-2 d-flex gap-2">
                <button type="submit" class="btn btn-custom-primary btn-lg flex-grow-1"><i class="fas fa-filter"></i></button>
                <a href="." class="btn btn-custom-primary btn-lg flex-grow-1"><i class="fas fa-undo"></i></a>
            </div>
        </form>
    </div>
</div>

<!-- Persons Table -->
<div class="card border-0 shadow">
    <div class="card-body p-0">
        <div class="table-responsive table-container">
            <table class="table table-striped table-hover mb-0">
                <thead class="table-fixed-header">
                    <tr>
                        <th>Revisar</th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=cedula&sort_direction={% if current_order == 'cedula' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Cedula
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=nombre_completo&sort_direction={% if current_order == 'nombre_completo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Nombre
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=cargo&sort_direction={% if current_order == 'cargo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Cargo
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=correo&sort_direction={% if current_order == 'correo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Correo
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=compania&sort_direction={% if current_order == 'compania' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Compania
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=estado&sort_direction={% if current_order == 'estado' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Estado
                            </a>
                        </th>
                        <th style="color: rgb(0, 0, 0);">Comentarios</th>
                        <th class="table-fixed-column" style="color: rgb(0, 0, 0);">Ver</th>
                    </tr>
                </thead>
                <tbody>
                    {% for person in persons %}
                        <tr>
                            <td>
                                {# Replace the existing <a> tag with a form that submits to your toggle view #}
                                <form action="{% url 'toggle_revisar_status' person.cedula %}" method="post" style="display:inline;">
                                    {% csrf_token %}
                                    <button type="submit"
                                            class="btn btn-link p-0 border-0 bg-transparent" {# Style button to look like a clickable icon #}
                                            title="{% if person.revisar %}Desmarcar para Revisar{% else %}Marcar para Revisar{% endif %}">
                                        <i class="fas fa-{% if person.revisar %}check-square text-warning{% else %}square text-secondary{% endif %}"
                                           style="padding-left: 20px; font-size: 1.25rem;"></i> {# Increased font-size for better clickability #}
                                    </button>
                                </form>
                            </td>
                            <td>{{ person.cedula }}</td>
                            <td>{{ person.nombre_completo }}</td>
                            <td>{{ person.cargo }}</td>
                            <td>{{ person.correo }}</td>
                            <td>{{ person.compania }}</td>
                            <td>
                                <span class="badge bg-{% if person.estado == 'Activo' %}success{% else %}danger{% endif %}">
                                    {{ person.estado }}
                                </span>
                            </td>
                            <td>{{ person.comments|truncatechars:30|default:"" }}</td>
                            <td class="table-fixed-column">
                                <a href="{% url 'person_details' person.cedula %}" 
                                   class="btn btn-custom-primary btn-sm"
                                   title="View details">
                                    <i class="bi bi-person-vcard-fill"></i>
                                </a>
                            </td>
                        </tr>
                    {% empty %}
                        <tr>
                            <td colspan="9" class="text-center py-4">
                                No hay personas marcadas para revisar.
                            </td>
                        </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
        
        <!-- Pagination -->
        {% if page_obj.has_other_pages %}
        <div class="p-3">
            <nav aria-label="Page navigation">
                <ul class="pagination justify-content-center">
                    {% if page_obj.has_previous %}
                        <li class="page-item">
                            <a class="page-link" href="?page=1{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="First">
                                <span aria-hidden="true">&laquo;&laquo;</span>
                            </a>
                        </li>
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.previous_page_number }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Previous">
                                <span aria-hidden="true">&laquo;</span>
                            </a>
                        </li>
                    {% endif %}
                    
                    {% for num in page_obj.paginator.page_range %}
                        {% if page_obj.number == num %}
                            <li class="page-item active"><a class="page-link" href="#">{{ num }}</a></li>
                        {% elif num > page_obj.number|add:'-3' and num < page_obj.number|add:'3' %}
                            <li class="page-item"><a class="page-link" href="?page={{ num }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}">{{ num }}</a></li>
                        {% endif %}
                    {% endfor %}
                    
                    {% if page_obj.has_next %}
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.next_page_number }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Next">
                                <span aria-hidden="true">&raquo;</span>
                            </a>
                        </li>
                        <li class="page-item">
                            <a class="page-link" href="?page={{ page_obj.paginator.num_pages }}{% for key, value in request.GET.items %}{% if key != 'page' %}&{{ key }}={{ value }}{% endif %}{% endfor %}" aria-label="Last">
                                <span aria-hidden="true">&raquo;&raquo;</span>
                            </a>
                        </li>
                    {% endif %}
                </ul>
            </nav>
        </div>
        {% endif %}
    </div>
</div>
{% endblock %}
"@ | Out-File -FilePath "core/templates/alerts.html" -Encoding utf8

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