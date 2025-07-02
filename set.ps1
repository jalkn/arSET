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
    python -m pip install django whitenoise django-bootstrap-v5 openpyxl pandas xlrd>=2.0.1 pdfplumber fitz msoffcrypto-tool fuzzywuzzy python-Levenshtein

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
    fecha_inicio = models.DateField(null=True, blank=True)
    q1 = models.BooleanField(default=False)
    q2 = models.BooleanField(default=False)
    q3 = models.BooleanField(default=False)
    q4 = models.BooleanField(default=False)
    q5 = models.BooleanField(default=False)
    q6 = models.BooleanField(default=False)
    q7 = models.BooleanField(default=False)
    q8 = models.BooleanField(default=False)
    q9 = models.BooleanField(default=False)
    q10 = models.BooleanField(default=False)
    q11 = models.BooleanField(default=False)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f"Conflictos para {self.person.nombre_completo} (ID: {self.id})"  
"@

# Create admin.py with enhanced configuration
Set-Content -Path "core/admin.py" -Value @" 
from django.contrib import admin
from django import forms
from django.utils.html import format_html
from django.urls import reverse
from core.models import Person, Conflict

class ConflictForm(forms.ModelForm):
    class Meta:
        model = Conflict
        fields = '__all__'
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Replace boolean field widgets with custom display
        for field_name in ['q1', 'q2', 'q3', 'q4', 'q5', 'q6', 'q7', 'q8', 'q9', 'q10', 'q11']:
            self.fields[field_name].widget = forms.Select(choices=[(True, 'YES'), (False, 'NO')])

@admin.register(Person)
class PersonAdmin(admin.ModelAdmin):
    list_display = ('cedula', 'nombre_completo', 'cargo', 'compania', 'estado', 'revisar')
    search_fields = ('cedula', 'nombre_completo', 'correo')
    list_filter = ('estado', 'compania', 'revisar')
    list_editable = ('revisar',)
    
    # Custom fields to show in detail view
    readonly_fields = ('cedula_with_actions', 'conflicts_link')
    
    fieldsets = (
        (None, {
            'fields': ('cedula_with_actions', 'nombre_completo', 'correo', 'estado', 'compania', 'cargo', 'revisar', 'comments')
        }),
        ('Related Records', {
            'fields': ('conflicts_link',),
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
    
    def get_fieldsets(self, request, obj=None):
        if obj is None:  # Add view
            return [(None, {'fields': ('cedula', 'nombre_completo', 'correo', 'estado', 'compania', 'cargo', 'revisar', 'comments')})]
        return super().get_fieldsets(request, obj)

@admin.register(Conflict)
class ConflictAdmin(admin.ModelAdmin):
    form = ConflictForm
    list_display = ('person', 'fecha_inicio') + tuple(
        f'get_{field}_display' for field in ['q1', 'q2', 'q3', 'q4', 'q5', 'q6', 'q7', 'q8', 'q9', 'q10', 'q11']
    )
    list_filter = ('q1', 'q2', 'q3', 'q4', 'q5', 'q6', 'q7', 'q8', 'q9', 'q10', 'q11')
    search_fields = ('person__nombre_completo', 'person__cedula')
    raw_id_fields = ('person',)

    fieldsets = (
        (None, {
            'fields': ('person', 'fecha_inicio')
        }),
        ('Conflict Questions', {
            'fields': ('q1', 'q2', 'q3', 'q4', 'q5', 'q6', 'q7', 'q8', 'q9', 'q10', 'q11'),
            'description': 'Answer "YES" or "NO" to each question'
        }),
    )

    def get_form(self, request, obj=None, **kwargs):
        form = super().get_form(request, obj, **kwargs)
        form.base_fields['q1'].label = 'Accionista de proveedor'
        form.base_fields['q2'].label = 'Familiar de accionista/empleado'
        form.base_fields['q3'].label = 'Accionista del grupo'
        form.base_fields['q4'].label = 'Actividades extralaborales'
        form.base_fields['q5'].label = 'Negocios con empleados'
        form.base_fields['q6'].label = 'Participacion en juntas'
        form.base_fields['q7'].label = 'Otro conflicto'
        form.base_fields['q8'].label = 'Conoce codigo de conducta'
        form.base_fields['q9'].label = 'Veracidad de informacion'
        form.base_fields['q10'].label = 'Familiar de funcionario'
        form.base_fields['q11'].label = 'Relacion con sector publico'
        return form

    # YES/NO display methods for list view
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
                   import_conflicts, conflict_list, import_persons, import_tcs, import_finances, person_details)

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
    path('conflicts/', views.conflict_list, name='conflict_list'),
    path('import-tcs/', views.import_tcs, name='import_tcs'),
    path('import-protected-excel/', views.import_finances, name='import_finances'),
    path('persons/<str:cedula>/', views.person_details, name='person_details'),
]
"@

# Update core/views.py with financial import
Set-Content -Path "core/views.py" -Value @"
import pandas as pd
from datetime import datetime
import os
from django.conf import settings
from django.http import HttpResponseRedirect
from django.shortcuts import render, redirect
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.contrib.auth import get_user_model
from django.contrib.auth.mixins import LoginRequiredMixin
from django.views.generic import TemplateView
from django.core.paginator import Paginator
from django.shortcuts import render
from core.models import Person, Conflict
from django.db.models import Q
import subprocess
import msoffcrypto
import io

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
        context['conflict_count'] = Conflict.objects.count()
        context['person_count'] = Person.objects.count()
        return context
    
@login_required
def main(request):
    return render(request, 'home.html')

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
                'nombre completo': 'nombre_completo',
                'correo': 'correo',
                'cedula': 'cedula',
                'estado': 'estado',
                'compania': 'compania',
                'cargo': 'cargo',
                'activo': 'activo' # Assuming 'activo' might be an input column for 'estado'
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
            
            # Convert 'nombre_completo' to title case if the column exists
            if 'nombre_completo' in df.columns:
                df['nombre_completo'] = df['nombre_completo'].str.title()
            else:
                messages.error(request, "Error: 'Nombre Completo' column not found in the Excel file.")
                return HttpResponseRedirect('/import/')

            # Define the columns for the output Excel file and ensure they are present
            output_columns_df = pd.DataFrame(columns=['NOMBRE COMPLETO', 'Cedula', 'Compania', 'CARGO'])
            
            # Populate the output DataFrame with data from the processed DataFrame
            if 'nombre_completo' in df.columns:
                output_columns_df['NOMBRE COMPLETO'] = df['nombre_completo']
            if 'cedula' in df.columns:
                output_columns_df['Cedula'] = df['cedula']
            if 'compania' in df.columns:
                output_columns_df['Compania'] = df['compania']
            if 'cargo' in df.columns:
                output_columns_df['CARGO'] = df['cargo']

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
                        'correo': row.get('correo', ''),
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
            dest_path = "core/src/conflictos.xlsx"
            with open(dest_path, 'wb+') as destination:
                for chunk in excel_file.chunks():
                    destination.write(chunk)
            
            subprocess.run(['python', 'core/conflicts.py'], check=True)
            
            import pandas as pd
            from core.models import Person, Conflict
            
            processed_file = "core/src/conflicts.xlsx"
            df = pd.read_excel(processed_file)
            df.columns = df.columns.str.lower().str.replace(' ', '_')
            
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
                    
                    Conflict.objects.update_or_create(
                        person=person,
                        defaults={
                            'fecha_inicio': row.get('fecha_de_inicio', None),
                            'q1': bool(row.get('q1', False)),
                            'q2': bool(row.get('q2', False)),
                            'q3': bool(row.get('q3', False)),
                            'q4': bool(row.get('q4', False)),
                            'q5': bool(row.get('q5', False)),
                            'q6': bool(row.get('q6', False)),
                            'q7': bool(row.get('q7', False)),
                            'q8': bool(row.get('q8', False)),
                            'q9': bool(row.get('q9', False)),
                            'q10': bool(row.get('q10', False)),
                            'q11': bool(row.get('q11', False))
                        }
                    )
                    
                except Exception as e:
                    messages.error(request, f"Error processing row {row}: {str(e)}")
                    continue
            
            messages.success(request, f'Archivo de conflictos importado exitosamente! {len(df)} registros procesados.')
        except Exception as e:
            messages.error(request, f'Error procesando archivo de conflictos: {str(e)}')
        
        return HttpResponseRedirect('/import/')
    
    return HttpResponseRedirect('/import/')

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
    }
    
    return render(request, 'persons.html', context)

@login_required
def conflict_list(request):
    search_query = request.GET.get('q', '')
    compania_filter = request.GET.get('compania', '')
    column_filter = request.GET.get('column', '')
    answer_filter = request.GET.get('answer', '')
    
    order_by = request.GET.get('order_by', 'person__nombre_completo')
    sort_direction = request.GET.get('sort_direction', 'asc')
    
    conflicts = Conflict.objects.select_related('person').all()
    
    if search_query:
        conflicts = conflicts.filter(
            Q(person__nombre_completo__icontains=search_query) |
            Q(person__cedula__icontains=search_query) |
            Q(person__correo__icontains=search_query))
    
    if compania_filter:
        conflicts = conflicts.filter(person__compania=compania_filter)
    
    if column_filter and answer_filter:
        filter_kwargs = {f"{column_filter}": answer_filter.lower() == 'yes'}
        conflicts = conflicts.filter(**filter_kwargs)
    
    if sort_direction == 'desc':
        order_by = f'-{order_by}'
    conflicts = conflicts.order_by(order_by)
    
    companias = Person.objects.exclude(compania='').values_list('compania', flat=True).distinct().order_by('compania')
    
    paginator = Paginator(conflicts, 25)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)
    
    context = {
        'conflicts': page_obj,
        'page_obj': page_obj,
        'companias': companias,
        'current_order': order_by.lstrip('-'),
        'current_direction': 'desc' if order_by.startswith('-') else 'asc',
        'all_params': {k: v for k, v in request.GET.items() if k not in ['page', 'order_by', 'sort_direction']},
    }
    
    return render(request, 'conflicts.html', context)

@login_required
def import_tcs(request):
    """View for importing credit card data from PDF files"""
    if request.method == 'POST' and request.FILES.getlist('visa_pdf_files'):
        pdf_files = request.FILES.getlist('visa_pdf_files')
        password = request.POST.get('visa_pdf_password', '')
        
        try:
            # Ensure visa directory exists
            visa_dir = os.path.join(settings.BASE_DIR, 'core', 'src', 'visa')
            os.makedirs(visa_dir, exist_ok=True)
            
            # Save password if provided
            if password:
                with open(os.path.join(visa_dir, 'password.txt'), 'w') as f:
                    f.write(password)
            
            # Save all PDF files
            for pdf_file in pdf_files:
                dest_path = os.path.join(visa_dir, pdf_file.name)
                with open(dest_path, 'wb+') as destination:
                    for chunk in pdf_file.chunks():
                        destination.write(chunk)
            
            # Process the PDFs
            subprocess.run(['python', 'core/tcs.py'], check=True, cwd=settings.BASE_DIR)
            
            # Count processed transactions
            output_file = os.path.join(visa_dir, f"VISA_{datetime.now().strftime('%Y%m%d')}.xlsx")
            if os.path.exists(output_file):
                df = pd.read_excel(output_file)
                record_count = len(df)
            else:
                record_count = 0
            
            messages.success(request, f'Archivos de tarjetas procesados exitosamente! {record_count} transacciones encontradas.')
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
        financial_reports = []  # Add your financial reports query here if needed
        
        context = {
            'myperson': person,
            'conflicts': conflicts,
            'financial_reports': financial_reports,
        }
        
        return render(request, 'details.html', context)
    except Person.DoesNotExist:
        messages.error(request, f"Person with ID {cedula} not found")
        return redirect('person_list')
"@

# Create core/conflicts.py
Set-Content -Path "core/conflicts.py" -Value @"
import pandas as pd
import os
from openpyxl.utils import get_column_letter
from openpyxl.styles import numbers

def extract_specific_columns(input_file, output_file, custom_headers=None):
    
    try:
        # Setup output directory
        os.makedirs(os.path.dirname(output_file), exist_ok=True)
        
        # Read raw data (no automatic parsing)
        df = pd.read_excel(input_file, header=None)
        
        # Column selection (first 11 + specified extras)
        base_cols = list(range(11))  # Columns 0-10 (A-K)
        extra_cols = [12,14,16,18,20,22,24,25,26,28]
        selected_cols = [col for col in base_cols + extra_cols if col < df.shape[1]]
        
        # Extract data with headers
        result = df.iloc[3:, selected_cols].copy()
        result.columns = df.iloc[2, selected_cols].values
        
        # Apply custom headers if provided
        if custom_headers is not None:
            if len(custom_headers) != len(result.columns):
                raise ValueError(f"Custom headers count ({len(custom_headers)}) doesn't match column count ({len(result.columns)})")
            result.columns = custom_headers
        
        # Merge C,D,E,F → C (indices 2,3,4,5)
        if all(c in selected_cols for c in [2,3,4,5]):
            result.iloc[:, 2] = result.iloc[:, 2:6].astype(str).apply(' '.join, axis=1)
            result.drop(result.columns[3:6], axis=1, inplace=True)
            selected_cols = [c for c in selected_cols if c not in [3,4,5]] 
            
        # Process "Nombre" column AFTER merging
        if "Nombre" in result.columns:
            # First replace actual NaN values with empty string
            result["Nombre"] = result["Nombre"].fillna("")
            # Then replace any "Nan" strings (case insensitive) with empty string
            result["Nombre"] = result["Nombre"].replace(r'(?i)\bNan\b', '', regex=True)
            # Clean up multiple spaces that might result from the replacement
            result["Nombre"] = result["Nombre"].str.replace(r'\s+', ' ', regex=True).str.strip()
            # Finally apply title case
            result["Nombre"] = result["Nombre"].str.title()
            
        # Special handling for Column J (input index 9)
        if 9 in selected_cols:
            j_pos = selected_cols.index(9)  # Find its position in output
            date_col = result.columns[j_pos]
            
            # Convert with European date format
            result[date_col] = pd.to_datetime(
                result[date_col],
                dayfirst=True,
                errors='coerce'
            )
            
            # Save with Excel formatting
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                result.to_excel(writer, index=False)
                
                # Get the worksheet and format the date column
                worksheet = writer.sheets['Sheet1']
                date_col_letter = get_column_letter(j_pos + 1)
                
                # Apply date format to all cells in the column
                for cell in worksheet[date_col_letter]:
                    if cell.row == 1:  # Skip header
                        continue
                    cell.number_format = 'DD/MM/YYYY'
                
                # Auto-adjust columns
                for idx, col in enumerate(result.columns):
                    col_letter = get_column_letter(idx+1)
                    worksheet.column_dimensions[col_letter].width = max(
                        len(str(col))+2,
                        result[col].astype(str).str.len().max()+2
                    )
        
        else:
            print("Warning: Column J not found in selected columns")
    
    except Exception as e:
        print(f"Error: {str(e)}")

# Example usage with custom headers
custom_headers = [
    "ID", "Cedula", "Nombre", "1er Nombre", "1er Apellido", 
    "2do Apellido", "Compañía", "Cargo", "Email", "Fecha de Inicio", 
    "Q1", "Q2", "Q3", "Q4", "Q5",
    "Q6", "Q7", "Q8", "Q9", "Q10", "Q11"
]

extract_specific_columns(
    input_file="core/src/conflictos.xlsx",
    output_file="core/src/conflicts.xlsx",
    custom_headers=custom_headers
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
    
    df[relative_col] = (
        df.groupby('Usuario')[column]
        .ffill()
        .pct_change(fill_method=None) * 100
    )
    
    df[relative_col] = df[relative_col].apply(lambda x: f"{x:.2f}%" if not pd.isna(x) else "0.00%")
    
    return df

def embed_trend_symbols(df, columns):
    """Add trend symbols to variation columns."""
    for col in columns:
        absolute_col = f'{col} Var. Abs.'
        relative_col = f'{col} Var. Rel.'
        
        if absolute_col in df.columns:
            df[absolute_col] = df.apply(
                lambda row: f"{row[absolute_col]:.2f} {get_trend_symbol(row[relative_col])}" 
                if pd.notna(row[absolute_col]) else "N/A ➡️",
                axis=1
            )
        
        if relative_col in df.columns:
            df[relative_col] = df.apply(
                lambda row: f"{row[relative_col]} {get_trend_symbol(row[relative_col])}", 
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
        BancoSaldo=('Banco - Saldo COP', 'sum'),
        Bienes=('Total Bienes', 'sum'),
        Inversiones=('Total Inversiones', 'sum')
    ).reset_index()

    for column in ['BancoSaldo', 'Bienes', 'Inversiones']:
        df_assets_grouped = calculate_variation(df_assets_grouped, column)
    
    df_assets_grouped = embed_trend_symbols(df_assets_grouped, ['BancoSaldo', 'Bienes', 'Inversiones'])
    return df_assets_grouped

def process_income_data(df_income):
    """Process income data with variations and trends."""
    df_income_grouped = df_income.groupby(['Usuario', 'Año Declaración']).agg(
        Ingresos=('Total Ingresos', 'sum'),
        Cant_Ingresos=('Cant_Ingresos', 'sum')
    ).reset_index()

    df_income_grouped = calculate_variation(df_income_grouped, 'Ingresos')
    df_income_grouped = embed_trend_symbols(df_income_grouped, ['Ingresos'])
    return df_income_grouped

def calculate_yearly_variations(df):
    """Calculate yearly variations for all columns."""
    df = df.sort_values(['Usuario', 'Año Declaración'])
    
    columns_to_analyze = [
        'Activos', 'Pasivos', 'Patrimonio', 
        'Apalancamiento', 'Endeudamiento',
        'BancoSaldo', 'Bienes', 'Inversiones', 'Ingresos',
        'Cant_Ingresos'
    ]
    
    new_columns = {}
    
    for column in [col for col in columns_to_analyze if col in df.columns]:
        grouped = df.groupby('Usuario')[column]
        
        for year in [2021, 2022, 2023, 2024]:
            abs_col = f'{year} {column} Var. Abs.'
            new_columns[abs_col] = grouped.diff()
            
            rel_col = f'{year} {column} Var. Rel.'
            pct_change = grouped.pct_change(fill_method=None) * 100
            new_columns[rel_col] = pct_change.apply(
                lambda x: f"{x:.2f}%" if not pd.isna(x) else "0.00%"
            )
    
    df = pd.concat([df, pd.DataFrame(new_columns)], axis=1)
    
    for column in [col for col in columns_to_analyze if col in df.columns]:
        for year in [2021, 2022, 2023, 2024]:
            abs_col = f'{year} {column} Var. Abs.'
            rel_col = f'{year} {column} Var. Rel.'
            
            if abs_col in df.columns:
                df[abs_col] = df.apply(
                    lambda row: f"{row[abs_col]:.2f} {get_trend_symbol(row[rel_col])}" 
                    if pd.notna(row[abs_col]) else "N/A ➡️",
                    axis=1
                )
            if rel_col in df.columns:
                df[rel_col] = df.apply(
                    lambda row: f"{row[rel_col]} {get_trend_symbol(row[rel_col])}", 
                    axis=1
                )
    
    return df

def calculate_sudden_wealth_increase(df):
    """Calculate sudden wealth increase rate (Aum. Pat. Subito) as decimal with 1 decimal place"""
    df = df.sort_values(['Usuario', 'Año Declaración'])
    
    # Calculate total wealth (Activo + Patrimonio)
    df['Capital'] = df['Activos'] + df['Patrimonio']
    
    # Calculate year-to-year change as decimal
    df['Aum. Pat. Subito'] = df.groupby('Usuario')['Capital'].pct_change(fill_method=None)
    
    # Format as decimal (1 place) with trend symbol
    df['Aum. Pat. Subito'] = df['Aum. Pat. Subito'].apply(
        lambda x: f"{x:.1f} {get_trend_symbol(f'{x*100:.1f}%')}" if not pd.isna(x) else "N/A ➡️"
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
        
        # Rename columns for output
        df_output.columns = [col.replace('Usuario', 'Cedula').replace('Compañía', 'Compania') 
                           for col in df_output.columns]
        
        # Ensure Cedula is string after renaming
        if 'Cedula' in df_output.columns:
            df_output['Cedula'] = df_output['Cedula'].astype(str)
        
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

# Read the Excel file
file_path = 'core/src/trends.xlsx'
df = pd.read_excel(file_path)

# Convert 'Cedula' column to string
df['Cedula'] = df['Cedula'].astype(str)

# Ensure all specified columns exist (create empty ones if they don't)
required_columns = [
    'Cedula', 'Nombre', 'Compania', 'Cargo', 'fkIdPeriodo', 'Año Declaración', 
    'Año Creación', 'Activos', 'Cant_Bienes', 'Cant_Bancos', 'Cant_Cuentas', 
    'Cant_Inversiones', 'Pasivos', 'Cant_Deudas', 'Patrimonio', 'Apalancamiento', 
    'Endeudamiento', 'Capital', 'Aum. Pat. Subito', 'Activos Var. Abs.', 
    'Activos Var. Rel.', 'Pasivos Var. Abs.', 'Pasivos Var. Rel.', 
    'Patrimonio Var. Abs.', 'Patrimonio Var. Rel.', 'Apalancamiento Var. Abs.', 
    'Apalancamiento Var. Rel.', 'Endeudamiento Var. Abs.', 'Endeudamiento Var. Rel.', 
    'BancoSaldo', 'Bienes', 'Inversiones', 'BancoSaldo Var. Abs.', 
    'BancoSaldo Var. Rel.', 'Bienes Var. Abs.', 'Bienes Var. Rel.', 
    'Inversiones Var. Abs.', 'Inversiones Var. Rel.', 'Ingresos', 
    'Cant_Ingresos', 'Ingresos Var. Abs.', 'Ingresos Var. Rel.'
]

# Add any missing columns with NaN values
for col in required_columns:
    if col not in df.columns:
        df[col] = None

# List of columns to convert to float (absolute variation columns)
float_columns = [
    'Activos Var. Abs.', 
    'Pasivos Var. Abs.', 
    'Patrimonio Var. Abs.', 
    'Apalancamiento Var. Abs.', 
    'Endeudamiento Var. Abs.',  
    'BancoSaldo Var. Abs.', 
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
    'BancoSaldo Var. Rel.', 
    'Bienes Var. Rel.', 
    'Inversiones Var. Rel.', 
    'Ingresos Var. Rel.'
]

# Convert absolute variation columns to float
for col in float_columns:
    if col in df.columns:
        df[col] = df[col].apply(lambda x: clean_and_convert(x, keep_trend=False))

# Process trend columns (handle infinity and preserve trend symbols)
for col in trend_columns:
    if col in df.columns:
        df[col] = df[col].apply(lambda x: clean_and_convert(x, keep_trend=True) 
                          if not pd.isna(x) and str(x).lower() not in ['inf', '-inf', 'inf%'] 
                          else np.nan)

# Special handling for 'Aum. Pat. Subito' column
if 'Aum. Pat. Subito' in df.columns:
    df['Aum. Pat. Subito'] = df['Aum. Pat. Subito'].apply(
        lambda x: np.nan if pd.isna(x) or "N/A ➡️" in str(x) else x
    )

# Reorder columns to match the specified order
df = df[required_columns]

# Read Personas.xlsx and merge with the current dataframe
personas_path = 'core/src/Personas.xlsx'
try:
    personas_df = pd.read_excel(personas_path)
    
    # Convert 'Cedula' to string in both dataframes to ensure matching
    personas_df['Cedula'] = personas_df['Cedula'].astype(str)
    df['Cedula'] = df['Cedula'].astype(str)
    
    # Rename columns in personas_df to match df column names
    personas_df = personas_df.rename(columns={
        'NOMBRE COMPLETO': 'Nombre',
        'CARGO': 'Cargo'
    })
    
    # Select only the columns we want to merge
    personas_merge = personas_df[['Cedula', 'Nombre', 'Compania', 'Cargo', 'Estado', 'Correo']]
    
    # Merge with the main dataframe on Cedula, keeping all records from df
    df = pd.merge(df, personas_merge, on='Cedula', how='left', suffixes=('', '_personas'))
    
    # For the merged columns, prioritize values from personas_df
    for col in ['Nombre', 'Compania', 'Cargo']:
        df[col] = df[f'{col}_personas'].fillna(df[col])
        df = df.drop(f'{col}_personas', axis=1)
        
except Exception as e:
    print(f"Error merging Personas.xlsx: {e}")

# Save the modified dataframe back to Excel
output_path = 'core/src/idTrends.xlsx'
df.to_excel(output_path, index=False)

print(f"File has been modified and saved as {output_path}")
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

.table-fixed-header {
    position: sticky;
    top: 0;
    z-index: 10;
    background-color: white;
}

.table-fixed-header th {
    position: sticky;
    top: 0;
    background-color: #f8f9fa; /* Match your table header color */
    z-index: 20;
}

/* Add a shadow to the fixed header */
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
    <a href="" class="btn btn-custom-primary">
        <i class="fas fa-chart-line" style="color: green;"></i>
    </a>
    <a href="" class="btn btn-custom-primary" title="Tarjetas">
        <i class="far fa-credit-card" style="color: blue;"></i>
    </a>
    <a href="{% url 'conflict_list' %}" class="btn btn-custom-primary">
        <i class="fas fa-balance-scale" style="color: orange;"></i>
    </a>
    <a href="" class="btn btn-custom-primary">
        <i class="fas fa-bell" style="color: red;"></i>
    </a>
    <a href="{% url 'import' %}" class="btn btn-custom-primary">
        <i class="fas fa-upload"></i> 
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
<div class="row justify-content-center">
    <div class="col-md-6">
        <div class="card border-0 shadow">
            <div class="card-body p-5">
                <div style="align-items: center; text-align: center;"> 
                    <h5 class="card-title">Bienvenido a ARPA</h5>
                    <p class="card-text">Automatizacion Robotica de Procesos de Auditoria</p>
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
    <a href="{% url 'person_list' %}" class="btn btn-custom-primary">
        <i class="fas fa-users"></i>
    </a>
    <a href="" class="btn btn-custom-primary">
        <i class="fas fa-chart-line" style="color: green;"></i>
    </a>
    <a href="" class="btn btn-custom-primary" title="Tarjetas">
        <i class="far fa-credit-card" style="color: blue;"></i>
    </a>
    <a href="{% url 'conflict_list' %}" class="btn btn-custom-primary">
        <i class="fas fa-balance-scale" style="color: orange;"></i>
    </a>
    <a href="" class="btn btn-custom-primary">
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
                        {{ protected_count }} Bienes y Rentas Registradas
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
                            <div class="form-text">Ingrese la clave si los PDFs están protegidos</div>
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

{% block title %}A R P A{% endblock %}
{% block navbar_title %}A R P A{% endblock %}

{% block navbar_buttons %}
<div>
    <a href="" class="btn btn-custom-primary" title="BienesyRentas">
        <i class="fas fa-chart-line" style="color: green;"></i>
    </a>
    <a href="" class="btn btn-custom-primary" title="Tarjetas">
        <i class="far fa-credit-card" style="color: blue;"></i>
    </a>
    <a href="{% url 'conflict_list' %}" class="btn btn-custom-primary">
        <i class="fas fa-balance-scale" style="color: orange;"></i>
    </a>
    <a href="" class="btn btn-custom-primary" title="Alertas">
        {% if alerts_count > 0 %}
            <span class="badge bg-danger">{{ alerts_count }}</span>
        {% endif %}
        {% if alerts_count == 0 %}
            <span class="badge bg-secondary">0</span>
        {% endif %}
        <i class="fas fa-bell" style="color: red;"></i>
    </a>
    <a href="{% url 'import' %}" class="btn btn-custom-primary">
        <i class="fas fa-upload"></i> 
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
                       placeholder="Buscar persona..." 
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
                                ID
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
@" 
{% extends "master.html" %}

{% block title %}Conflictos de Interes{% endblock %}
{% block navbar_title %}Conflictos de Interes{% endblock %}

{% block navbar_buttons %}
<div>
    <a href="{% url 'person_list' %}" class="btn btn-custom-primary">
        <i class="fas fa-users"></i>
    </a>
    <a href="" class="btn btn-custom-primary">
        <i class="fas fa-chart-line" style="color: green;"></i>
    </a>
    <a href="" class="btn btn-custom-primary" title="Tarjetas">
        <i class="far fa-credit-card" style="color: blue;"></i>
    </a>
    <a href="" class="btn btn-custom-primary">
        <i class="fas fa-bell" style="color: red;"></i>
    </a>
    <a href="{% url 'import' %}" class="btn btn-custom-primary" title="Importar">
        <i class="fas fa-upload"></i>
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
        <form method="get" action="." class="row g-3 align-items-center no-loading">
            <div class="d-flex align-items-center">
                <span class="badge bg-success">
                    {{ page_obj.paginator.count }} registros
                </span>
                {% if request.GET.q or request.GET.compania or request.GET.column or request.GET.answer %}
                {% endif %}
            </div>
            <!-- General Search -->
            <div class="col-md-4">
                <input type="text" 
                       name="q" 
                       class="form-control form-control-lg" 
                       placeholder="Buscar..." 
                       value="{{ request.GET.q }}">
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
            
            <!-- Column Selector -->
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
            
            <!-- Answer Selector -->
            <div class="col-md-2">
                <select name="answer" class="form-select form-select-lg">
                    <option value="">Selecciona Respuesta</option>
                    <option value="yes" {% if request.GET.answer == 'yes' %}selected{% endif %}>Si</option>
                    <option value="no" {% if request.GET.answer == 'no' %}selected{% endif %}>No</option>
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

<!-- Conflicts Table -->
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
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=q1&sort_direction={% if current_order == 'q1' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Accionista de proveedor
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=q2&sort_direction={% if current_order == 'q2' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Familiar de accionista/empleado
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=q3&sort_direction={% if current_order == 'q3' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Accionista del grupo
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=q4&sort_direction={% if current_order == 'q4' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Actividades extralaborales
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=q5&sort_direction={% if current_order == 'q5' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Negocios con empleados
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=q6&sort_direction={% if current_order == 'q6' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Participacion en juntas
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=q7&sort_direction={% if current_order == 'q7' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Otro conflicto
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=q8&sort_direction={% if current_order == 'q8' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Conoce codigo de conducta
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=q9&sort_direction={% if current_order == 'q9' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Veracidad de informacion
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=q10&sort_direction={% if current_order == 'q10' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Familiar de funcionario
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=q11&sort_direction={% if current_order == 'q11' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Relacion con sector publico
                            </a>
                        </th>
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
                            <td class="text-center">{% if conflict.q1 %}<i style="color: red;">SI</i>{% else %}<i style="color: green;">NO</i>{% endif %}</td>
                            <td class="text-center">{% if conflict.q2 %}<i style="color: red;">SI</i>{% else %}<i style="color: green;">NO</i>{% endif %}</td>
                            <td class="text-center">{% if conflict.q3 %}<i style="color: red;">SI</i>{% else %}<i style="color: green;">NO</i>{% endif %}</td>
                            <td class="text-center">{% if conflict.q4 %}<i style="color: red;">SI</i>{% else %}<i style="color: green;">NO</i>{% endif %}</td>
                            <td class="text-center">{% if conflict.q5 %}<i style="color: red;">SI</i>{% else %}<i style="color: green;">NO</i>{% endif %}</td>
                            <td class="text-center">{% if conflict.q6 %}<i style="color: red;">SI</i>{% else %}<i style="color: green;">NO</i>{% endif %}</td>
                            <td class="text-center">{% if conflict.q7 %}<i style="color: red;">SI</i>{% else %}<i style="color: green;">NO</i>{% endif %}</td>
                            <td class="text-center">{% if conflict.q8 %}<i style="color: green;">SI</i>{% else %}<i style="color: red;">NO</i>{% endif %}</td>
                            <td class="text-center">{% if conflict.q9 %}<i style="color: green;">SI</i>{% else %}<i style="color: RED;">NO</i>{% endif %}</td>
                            <td class="text-center">{% if conflict.q10 %}<i style="color: red;">SI</i>{% else %}<i style="color: green;">NO</i>{% endif %}</td>
                            <td class="text-center">{% if conflict.q11 %}<i style="color: red;">SI</i>{% else %}<i style="color: green;">NO</i>{% endif %}</td>
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
                            <td colspan="15" class="text-center py-4">
                                {% if request.GET.q or request.GET.compania or request.GET.column or request.GET.answer %}
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
"@ | Out-File -FilePath "core/templates/conflicts.html" -Encoding utf8

@"
{% extends "master.html" %}

{% block title %}Tarjetas de Credito{% endblock %}
{% block navbar_title %}Tarjetas de Credito{% endblock %}

{% block navbar_buttons %}
<div>
    <a href="" class="btn btn-custom-primary">
        <i class="fas fa-users"></i>
    </a>
    <a href="{% url 'financial_list' %}" class="btn btn-custom-primary">
        <i class="fas fa-chart-line" style="color: green;"></i>
    </a>
    <a href="" class="btn btn-custom-primary">
        <i class="fas fa-balance-scale" style="color: orange;"></i>
    </a>
    <a href="" class="btn btn-custom-primary">
        <i class="fas fa-bell" style="color: red;"></i>
    </a>
    <a href="" class="btn btn-custom-primary" title="Importar">
        <i class="fas fa-upload"></i>
    </a>
    <a href="?{% for key, value in request.GET.items %}{{ key }}={{ value }}&{% endfor %}export=excel" class="btn btn-custom-primary btn-my-green" title="Exportar">
        <i class="fas fa-file-excel"></i>
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
<!-- Search Form -->
<div class="card mb-4 border-0 shadow" style="background-color:rgb(224, 224, 224);">
    <div class="card-body">
        <form method="get" action="." class="row g-3 align-items-center">
            <!-- General Search 
            <div class="col-md-4">
                <input type="text" 
                       name="q" 
                       class="form-control form-control-lg" 
                       placeholder="Buscar tarjetahabiente..." 
                       value="{{ request.GET.q }}">
            </div> -->
            
            <!-- Card Type Filter -->
            <div class="col-md-3">
                <select name="card_type" class="form-select form-select-lg">
                    <option value="">Tipo</option>
                    <option value="MC" {% if request.GET.card_type == 'MC' %}selected{% endif %}>Mastercard</option>
                    <option value="VI" {% if request.GET.card_type == 'VI' %}selected{% endif %}>Visa</option>
                </select>
            </div>
            
            <!-- Date Range -->
            <div class="col-md-3">
                <input type="date" 
                       name="date_from" 
                       class="form-control form-control-lg" 
                       value="{{ request.GET.date_from }}">
            </div>
            <div class="col-md-3">
                <input type="date" 
                       name="date_to" 
                       class="form-control form-control-lg" 
                       value="{{ request.GET.date_to }}">
            </div>
            
            <!-- Submit Buttons -->
            <div class="col-md-2 d-flex gap-2">
                <button type="submit" class="btn btn-custom-primary btn-lg flex-grow-1"><i class="fas fa-filter"></i></button>
                <a href="." class="btn btn-custom-primary btn-lg flex-grow-1"><i class="fas fa-undo"></i></a>
            </div>
        </form>
    </div>
</div>

<!-- Cards Table -->
<div class="card border-0 shadow">
    <div class="card-body p-0">
        <div class="table-responsive table-container">
            <table class="table table-striped table-hover mb-0">
                <thead class="table-fixed-header">
                    <tr>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=person__nombre_completo&sort_direction={% if current_order == 'person__nombre_completo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Tarjetahabiente
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=card_type&sort_direction={% if current_order == 'card_type' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Tipo
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=card_number&sort_direction={% if current_order == 'card_number' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Numero
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=transaction_date&sort_direction={% if current_order == 'transaction_date' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Fecha
                            </a>
                        </th>
                        <th style="color: rgb(0, 0, 0);">DescripciÃ³n</th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=original_value&sort_direction={% if current_order == 'original_value' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Valor Original
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=exchange_rate&sort_direction={% if current_order == 'exchange_rate' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Tasa
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=charges&sort_direction={% if current_order == 'charges' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Cargos
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=balance&sort_direction={% if current_order == 'balance' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Saldo
                            </a>
                        </th>
                        <th style="color: rgb(0, 0, 0);">Cuotas</th>
                        <th style="color: rgb(0, 0, 0);">Archivo</th>
                        <th class="table-fixed-column" style="color: rgb(0, 0, 0);">Ver</th>
                    </tr>
                </thead>
                <!-- In your table body -->
                <tbody>
                    {% for card in page_obj.object_list %}  <!-- Changed from cards to page_obj.object_list -->
                    <tr {% if card.person.revisar %}class="table-warning"{% endif %}>
                        <td>{{ card.person.nombre_completo }}</td>
                        <td>{{ card.get_card_type_display }}</td>
                        <td>**** **** **** {{ card.card_number|slice:"-4:" }}</td>
                        <td>{{ card.transaction_date|date:"d/m/Y" }}</td>
                        <td>{{ card.description|truncatechars:30 }}</td>
                        <td>`$`{{ card.original_value|floatformat:2 }}</td>
                        <td>{{ card.exchange_rate|default_if_none:"-"|floatformat:4 }}</td>
                        <td>`$`{{ card.charges|default_if_none:"-"|floatformat:2 }}</td>
                        <td>`$`{{ card.balance|default_if_none:"-"|floatformat:2 }}</td>
                        <td>{{ card.installments|default:"-" }}</td>
                        <td>{{ card.source_file|truncatechars:15 }}</td>
                        <td class="table-fixed-column">
                            <a href="/persons/details/{{ card.person.cedula }}/" 
                            class="btn btn-custom-primary btn-sm"
                            title="Ver detalles">
                                <i class="bi bi-person-vcard-fill"></i>
                            </a>
                        </td>
                    </tr>
                    {% empty %}
                        <tr>
                            <td colspan="12" class="text-center py-4">
                                {% if request.GET.q or request.GET.card_type or request.GET.date_from or request.GET.date_to %}
                                    No se encontraron transacciones con los filtros aplicados.
                                {% else %}
                                    No hay transacciones de tarjetas registradas.
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
"@ | Out-File -FilePath "core/templates/cards.html" -Encoding utf8

# finances template
@" 
{% extends "master.html" %}

{% block title %}Bienes y Rentas{% endblock %}
{% block navbar_title %}Bienes y Rentas{% endblock %}

{% block navbar_buttons %}
<div>
    <a href="{% url 'person_list' %}" class="btn btn-custom-primary">
        <i class="fas fa-users"></i>
    </a>
    <a href="" class="btn btn-custom-primary" title="Tarjetas">
        <i class="far fa-credit-card" style="color: blue;"></i>
    </a>
    <a href="{% url 'conflict_list' %}" class="btn btn-custom-primary">
        <i class="fas fa-balance-scale" style="color: orange;"></i>
    </a>
    <a href="" class="btn btn-custom-primary">
        <i class="fas fa-bell" style="color: red;"></i>
    </a>
    <a href="{% url 'import' %}" class="btn btn-custom-primary" title="Importar">
        <i class="fas fa-upload"></i>
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
<!-- Search Form -->
<div class="card mb-4 border-0 shadow" style="background-color:rgb(224, 224, 224);">
    <div class="card-body">
        <form method="get" action="." class="row g-3 align-items-center">
            <!-- General Search -->
            <div class="col-md-3">
                <input type="text" 
                       name="q" 
                       class="form-control form-control-lg" 
                       placeholder="Buscar persona..." 
                       value="{{ request.GET.q }}">
            </div>
            
            <!-- Column Selector -->
            <div class="col-md-3">
                <select name="column" class="form-select form-select-lg">
                    <option value="">Selecciona Columna</option>
                    <option value="ano_declaracion" {% if request.GET.column == 'ano_declaracion' %}selected{% endif %}>Ano Declaracion</option>
                    <option value="aum_pat_subito" {% if request.GET.column == 'aum_pat_subito' %}selected{% endif %}>Aum. Pat. Subito</option>
                    <option value="activos_var_rel" {% if request.GET.column == 'activos_var_rel' %}selected{% endif %}>Activos Var. Rel.</option>
                    <option value="pasivos_var_rel" {% if request.GET.column == 'pasivos_var_rel' %}selected{% endif %}>Pasivos Var. Rel.</option>
                    <option value="patrimonio_var_rel" {% if request.GET.column == 'patrimonio_var_rel' %}selected{% endif %}>Patrimonio Var. Rel.</option>
                    <option value="apalancamiento_var_rel" {% if request.GET.column == 'apalancamiento_var_rel' %}selected{% endif %}>Apalancamiento Var. Rel.</option>
                    <option value="endeudamiento_var_rel" {% if request.GET.column == 'endeudamiento_var_rel' %}selected{% endif %}>Endeudamiento Var. Rel.</option>
                    <option value="banco_saldo_var_rel" {% if request.GET.column == 'banco_saldo_var_rel' %}selected{% endif %}>BancoSaldo Var. Rel.</option>
                    <option value="bienes_var_rel" {% if request.GET.column == 'bienes_var_rel' %}selected{% endif %}>Bienes Var. Rel.</option>
                    <option value="inversiones_var_rel" {% if request.GET.column == 'inversiones_var_rel' %}selected{% endif %}>Inversiones Var. Rel.</option>
                </select>
            </div>
            
            <!-- Operator Selector -->
            <div class="col-md-2">
                <select name="operator" class="form-select form-select-lg">
                    <option value="">Selecciona operador</option>
                    <option value=">" {% if request.GET.operator == '>' %}selected{% endif %}>Mayor que</option>
                    <option value="<" {% if request.GET.operator == '<' %}selected{% endif %}>Menor que</option>
                    <option value="=" {% if request.GET.operator == '=' %}selected{% endif %}>Igual a</option>
                    <option value=">=" {% if request.GET.operator == '>=' %}selected{% endif %}>Mayor o igual</option>
                    <option value="<=" {% if request.GET.operator == '<=' %}selected{% endif %}>Menor o igual</option>
                    <option value="between" {% if request.GET.operator == 'between' %}selected{% endif %}>Entre</option>
                    <option value="contains" {% if request.GET.operator == 'contains' %}selected{% endif %}>Contiene</option>
                </select>
            </div>
            
            <!-- Value Input -->
            <div class="col-md-2">
                <input type="text" 
                       name="value" 
                       class="form-control form-control-lg" 
                       placeholder="Valor" 
                       value="{{ request.GET.value }}">
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
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=nombre_completo&sort_direction={% if current_order == 'nombre_completo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Nombre
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=compania&sort_direction={% if current_order == 'compania' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Compania
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=cargo&sort_direction={% if current_order == 'cargo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Cargo
                            </a>
                        </th>
                        <th style="color: rgb(0, 0, 0);">Comentarios</th>
                        <th style="color: rgb(0, 0, 0);">Periodo</th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__ano_declaracion&sort_direction={% if current_order == 'financial_reports__ano_declaracion' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Ano
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__aum_pat_subito&sort_direction={% if current_order == 'financial_reports__aum_pat_subito' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Aum. Pat. Subito
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__patrimonio&sort_direction={% if current_order == 'financial_reports__patrimonio' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Patrimonio
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__patrimonio_var_rel&sort_direction={% if current_order == 'financial_reports__patrimonio_var_rel' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Patrimonio Var. Rel. % 
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__patrimonio_var_abs&sort_direction={% if current_order == 'financial_reports__patrimonio_var_abs' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Patrimonio Var. Abs. $
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__activos&sort_direction={% if current_order == 'financial_reports__activos' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Activos
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__activos_var_rel&sort_direction={% if current_order == 'financial_reports__activos_var_rel' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Activos Var. Rel. %
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__activos_var_abs&sort_direction={% if current_order == 'financial_reports__activos_var_abs' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Activos Var. Abs. $
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__pasivos&sort_direction={% if current_order == 'financial_reports__pasivos' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Pasivos
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__pasivos_var_rel&sort_direction={% if current_order == 'financial_reports__pasivos_var_rel' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Pasivos Var. Rel. 
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__pasivos_var_abs&sort_direction={% if current_order == 'financial_reports__pasivos_var_abs' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Pasivos Var. Abs. $
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__cant_deudas&sort_direction={% if current_order == 'financial_reports__cant_deudas' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Cant. Deudas
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__ingresos&sort_direction={% if current_order == 'financial_reports__ingresos' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Ingresos
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__ingresos_var_rel&sort_direction={% if current_order == 'financial_reports__ingresos_var_rel' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Ingresos Var. Rel. %
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__ingresos_var_abs&sort_direction={% if current_order == 'financial_reports__ingresos_var_abs' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Ingresos Var. Abs. $
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__cant_ingresos&sort_direction={% if current_order == 'financial_reports__cant_ingresos' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Cant. Ingresos
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__banco_saldo&sort_direction={% if current_order == 'financial_reports__banco_saldo' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Bancos Saldo
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__banco_saldo_var_rel&sort_direction={% if current_order == 'financial_reports__banco_saldo_var_rel' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Bancos Var. %
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__banco_saldo_var_abs&sort_direction={% if current_order == 'financial_reports__banco_saldo_var_abs' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Bancos Var. $
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__cant_cuentas&sort_direction={% if current_order == 'financial_reports__cant_cuentas' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Cant. Cuentas
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__cant_bancos&sort_direction={% if current_order == 'financial_reports__cant_bancos' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Cant. Bancos
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__bienes_valor&sort_direction={% if current_order == 'financial_reports__bienes_valor' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Bienes Valor
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__bienes_var_rel&sort_direction={% if current_order == 'financial_reports__bienes_var_rel' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Bienes Var. %
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__bienes_var_abs&sort_direction={% if current_order == 'financial_reports__bienes_var_abs' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Bienes Var. $
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__cant_bienes&sort_direction={% if current_order == 'financial_reports__cant_bienes' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Cant. Bienes
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__inversiones_valor&sort_direction={% if current_order == 'financial_reports__inversiones_valor' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Inversiones Valor
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__inversiones_var_rel&sort_direction={% if current_order == 'financial_reports__inversiones_var_rel' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Inversiones Var. %
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__inversiones_var_abs&sort_direction={% if current_order == 'financial_reports__inversiones_var_abs' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Inversiones Var. $
                            </a>
                        </th>
                        <th>
                            <a href="?{% for key, value in all_params.items %}{{ key }}={{ value }}&{% endfor %}order_by=financial_reports__cant_inversiones&sort_direction={% if current_order == 'financial_reports__cant_inversiones' and current_direction == 'asc' %}desc{% else %}asc{% endif %}" style="text-decoration: none; color: rgb(0, 0, 0);">
                                Cant. Inversiones
                            </a>
                        </th>
                        <th class="table-fixed-column" style="color: rgb(0, 0, 0);">Ver</th>
                    </tr>
                </thead>
                <tbody>
                    {% for person in persons %}
                        {% for report in person.financial_reports.all %}
                        <tr {% if person.revisar %}class="table-warning"{% endif %}>
                            <td>
                                <a href="/admin/core/person/{{ person.cedula }}/change/" style="text-decoration: none;" title="{% if person.revisar %}Marcado para revisar{% else %}No marcado{% endif %}">
                                    <i class="fas fa-{% if person.revisar %}check-square text-warning{% else %}square text-secondary{% endif %}" style="padding-left: 20px;"></i>
                                </a>
                            </td>
                            <td>{{ person.nombre_completo }}</td>
                            <td>{{ person.compania }}</td>
                            <td>{{ person.cargo }}</td>
                            <td>{{ person.comments|truncatechars:30|default:"" }}</td>
                            <td>{{ report.fkIdPeriodo|floatformat:"0"|default:"-" }}</td>
                            <td>{{ report.ano_declaracion|floatformat:"0"|default:"-" }}</td>
                            <td>{{ report.aum_pat_subito|default:"-" }}</td>
                            <td>{{ report.patrimonio|default:"-" }}</td>
                            <td>{{ report.patrimonio_var_rel|default:"-" }}</td>
                            <td>{{ report.patrimonio_var_abs|default:"-" }}</td>
                            <td>{{ report.activos|default:"-" }}</td>
                            <td>{{ report.activos_var_rel|default:"-" }}</td>
                            <td>{{ report.activos_var_abs|default:"-" }}</td>
                            <td>{{ report.pasivos|default:"-" }}</td>
                            <td>{{ report.pasivos_var_rel|default:"-" }}</td>
                            <td>{{ report.pasivos_var_abs|default:"-" }}</td>
                            <td>{{ report.cant_deudas|default:"-" }}</td>
                            <td>{{ report.ingresos|default:"-" }}</td>
                            <td>{{ report.ingresos_var_rel|default:"-" }}</td>
                            <td>{{ report.ingresos_var_abs|default:"-" }}</td>
                            <td>{{ report.cant_ingresos|default:"-" }}</td>
                            <td>{{ report.banco_saldo|default:"-" }}</td>
                            <td>{{ report.banco_saldo_var_rel|default:"-" }}</td>
                            <td>{{ report.banco_saldo_var_abs|default:"-" }}</td>
                            <td>{{ report.cant_cuentas|default:"-" }}</td>
                            <td>{{ report.cant_bancos|default:"-" }}</td>
                            <td>{{ report.bienes|default:"-" }}</td>
                            <td>{{ report.bienes_var_rel|default:"-" }}</td>
                            <td>{{ report.bienes_var_abs|default:"-" }}</td>
                            <td>{{ report.cant_bienes|default:"-" }}</td>
                            <td>{{ report.inversiones|default:"-" }}</td>
                            <td>{{ report.inversiones_var_rel|default:"-" }}</td>
                            <td>{{ report.inversiones_var_abs|default:"-" }}</td>
                            <td>{{ report.cant_inversiones|default:"-" }}</td>
                            <td class="table-fixed-column">
                                <a href="/persons/details/{{ person.cedula }}/" 
                                   class="btn btn-custom-primary btn-sm"
                                   title="View details">
                                    <i class="bi bi-person-vcard-fill"></i>
                                </a>
                            </td>
                        </tr>
                        {% empty %}
                        <tr>
                            <td colspan="14">{{ person.nombre_completo }} - No hay reportes financieros</td>
                        </tr>
                        {% endfor %}
                    {% empty %}
                        <tr>
                            <td colspan="14" class="text-center py-4">
                                {% if request.GET.q or request.GET.column or request.GET.operator or request.GET.value %}
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
"@ | Out-File -FilePath "core/templates/finances.html" -Encoding utf8

# details template
@" 
{% extends "master.html" %}
{% load humanize %}

{% block title %}Detalles - {{ myperson.nombre_completo }}{% endblock %}
{% block navbar_title %}{{ myperson.nombre_completo }}{% endblock %}

{% block navbar_buttons %}
<a href="/admin/core/person/{{ myperson.cedula }}/change/" class="btn btn-outline-dark" title="Admin">
    <i class="fas fa-wrench"></i>
</a>
<a href="/" class="btn btn-custom-primary"><i class="fas fa-arrow-right"></i></a>
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
            <div class="card-header bg-light">
                <h5 class="mb-0">Conflictos Declarados</h5>
            </div>
            <div class="card-body p-0">
                {% if conflicts %}
                {% for conflict in conflicts %}
                <div class="table-responsive">
                    <table class="table table-striped table-hover mb-0">
                        <tbody>
                            <tr>
                                <th scope="row">Fecha de Inicio</th>
                                <td>{{ conflict.fecha_inicio|date:"d/m/Y"|default:"-" }}</td>
                            </tr>
                            <tr>
                                <th scope="row">Accionista de algun proveedor del grupo</th>
                                <td class="text-center">{% if conflict.q1 %}<i style="color: red;">SI</i>{% else %}<i style="color: green;">NO</i>{% endif %}</td>
                            </tr>
                            <tr>
                                <th scope="row">Familiar de algun accionista, proveedor o empleado</th>
                                <td class="text-center">{% if conflict.q2 %}<i style="color: red;">SI</i>{% else %}<i style="color: green;">NO</i>{% endif %}</td>
                            </tr>
                            <tr>
                                <th scope="row">Accionista de alguna compania del grupo</th>
                                <td class="text-center">{% if conflict.q3 %}<i style="color: red;">SI</i>{% else %}<i style="color: green;">NO</i>{% endif %}</td>
                            </tr>
                            <tr>
                                <th scope="row">Actividades extralaborales</th>
                                <td class="text-center">{% if conflict.q4 %}<i style="color: red;">SI</i>{% else %}<i style="color: green;">NO</i>{% endif %}</td>
                            </tr>
                            <tr>
                                <th scope="row">Negocios o bienes con empleados del grupo</th>
                                <td class="text-center">{% if conflict.q5 %}<i style="color: red;">SI</i>{% else %}<i style="color: green;">NO</i>{% endif %}</td>
                            </tr>
                            <tr>
                                <th scope="row">Participacion en juntas o consejos directivos</th>
                                <td class="text-center">{% if conflict.q6 %}<i style="color: red;">SI</i>{% else %}<i style="color: green;">NO</i>{% endif %}</td>
                            </tr>
                            <tr>
                                <th scope="row">Potencial conflicto diferente a los anteriores</th>
                                <td class="text-center">{% if conflict.q7 %}<i style="color: red;">SI</i>{% else %}<i style="color: green;">NO</i>{% endif %}</td>
                            </tr>
                            <tr>
                                <th scope="row">Consciente del codigo de conducta empresarial</th>
                                <td class="text-center">{% if conflict.q8 %}<i style="color: green;">SI</i>{% else %}<i style="color: red;">NO</i>{% endif %}</td>
                            </tr>
                            <tr>
                                <th scope="row">Veracidad de la informacion consignada</th>
                                <td class="text-center">{% if conflict.q9 %}<i style="color: green;">SI</i>{% else %}<i style="color: RED;">NO</i>{% endif %}</td>
                            </tr>
                            <tr>
                                <th scope="row">Familiar de algun funcionario publico</th>
                                <td class="text-center">{% if conflict.q10 %}<i style="color: red;">SI</i>{% else %}<i style="color: green;">NO</i>{% endif %}</td>
                            </tr>
                            <tr>
                                <th scope="row">Relacion con el sector publico o funcionario publico</th>
                                <td class="text-center">{% if conflict.q11 %}<i style="color: red;">SI</i>{% else %}<i style="color: green;">NO</i>{% endif %}</td>
                            </tr>
                        </tbody>
                    </table>
                </div>
                <hr> {% endfor %}
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
                                <th>Apalancamiento</th>
                                <th>Endeudamiento</th>
                                <th>Indice</th>
                            </tr>
                        </thead>
                        <tbody>
                            {% for report in financial_reports %}
                            <tr>
                                <td>{{ report.ano_declaracion|floatformat:"0"|default:"-" }}</td>
                                <th>Relativa</th>
                                <td>{{ report.activos_var_rel|default:"-" }}</td>
                                <td>{{ report.pasivos_var_rel|default:"-" }}</td>
                                <td>{{ report.ingresos_var_rel|default:"-" }}</td>
                                <td>{{ report.patrimonio_var_rel|default:"-" }}</td>
                                <td>{{ report.banco_saldo_var_rel|default:"-" }}</td>
                                <td>{{ report.bienes_var_rel|default:"-" }}</td>
                                <td>{{ report.inversiones_var_rel|default:"-" }}</td>
                                <td>{{ report.apalancamiento_var_rel|default:"-" }}</td>
                                <td>{{ report.endeudamiento_var_rel|default:"-" }}</td>
                                <td>{{ report.aum_pat_subito|default:"-" }}</td>
                            </tr>
                            <tr>
                                <th></th>
                                <th scope="col">Absoluta</th>
                                <td>{{ report.activos_var_abs|intcomma|default:"-" }}</td>
                                <td>{{ report.pasivos_var_abs|intcomma|default:"-" }}</td>
                                <td>{{ report.ingresos_var_abs|intcomma|default:"-" }}</td>
                                <td>{{ report.patrimonio_var_abs|intcomma|default:"-" }}</td>
                                <td>{{ report.banco_saldo_var_abs|intcomma|default:"-" }}</td>
                                <td>{{ report.bienes_var_abs|intcomma|default:"-" }}</td>
                                <td>{{ report.inversiones_var_abs|intcomma|default:"-" }}</td>
                                <td>{{ report.apalancamiento_var_abs|default:"-" }}</td>
                                <td>{{ report.endeudamiento_var_abs|default:"-" }}</td>
                                <td>{{ report.capital_var_abs|intcomma|default:"-" }}</td>
                            </tr>
                            <tr>
                                <td></td>
                                <th scope="col">Total</th>
                                <td>&#36;{{ report.activos|floatformat:2|intcomma|default:"-" }}</td>
                                <td>&#36;{{ report.pasivos|floatformat:2|intcomma|default:"-" }}</td>
                                <td>&#36;{{ report.ingresos|floatformat:2|intcomma|default:"-" }}</td>
                                <td>&#36;{{ report.patrimonio|floatformat:2|intcomma|default:"-" }}</td>
                                <td>&#36;{{ report.banco_saldo|floatformat:2|intcomma|default:"-" }}</td>
                                <td>&#36;{{ report.bienes|floatformat:2|intcomma|default:"-" }}</td>
                                <td>&#36;{{ report.inversiones|floatformat:2|intcomma|default:"-" }}</td>
                                <td>{{ report.apalancamiento|floatformat:2|default:"-" }}</td>
                                <td>{{ report.endeudamiento|floatformat:2|default:"-" }}</td>
                                <td>&#36;{{ report.capital|floatformat:2|intcomma|default:"-" }}</td>
                            </tr>
                            <tr>
                                <th></th>
                                <th scope="col">Cant.</th>
                                <td></td>
                                <td>{{ report.cant_deudas|default:"-" }}</td>
                                <td>{{ report.cant_ingresos|default:"-" }}</td>
                                <td></td>
                                <td>C{{ report.cant_cuentas|default:"-" }} B{{ report.cant_bancos|default:"-" }}</td>
                                <td>{{ report.cant_bienes|default:"-" }}</td>
                                <td>{{ report.cant_inversiones|default:"-" }}</td>
                                <td></td>
                                <td></td>
                                <td></td>
                            </tr>
                                
                            </tr>
                            {% empty %}
                            <tr>
                                <td colspan="8" class="text-center py-4">
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
{% endblock %}
"@ | Out-File -FilePath "core/templates/details.html" -Encoding utf8

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