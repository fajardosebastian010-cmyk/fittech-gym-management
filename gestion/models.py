from django.db import models
from django.contrib.auth.models import AbstractBaseUser, BaseUserManager, PermissionsMixin
from django.utils import timezone
from datetime import timedelta

class UsuarioManager(BaseUserManager):
    def create_user(self, correo, password=None, **extra_fields):
        if not correo:
            raise ValueError('El correo es obligatorio')
        correo = self.normalize_email(correo)
        user = self.model(correo=correo, **extra_fields)
        user.set_password(password)
        user.save(using=self._db)
        return user

    def create_superuser(self, correo, password=None, **extra_fields):
        extra_fields.setdefault('rol', 'administrador')
        extra_fields.setdefault('is_staff', True)
        extra_fields.setdefault('is_superuser', True)
        return self.create_user(correo, password, **extra_fields)

class Usuario(AbstractBaseUser, PermissionsMixin):
    ROLES = [
        ('empleado', 'Empleado'),
        ('administrador', 'Administrador'),
    ]
    
    nombre = models.CharField(max_length=100)
    correo = models.EmailField(unique=True)
    rol = models.CharField(max_length=20, choices=ROLES)
    is_active = models.BooleanField(default=True)
    is_staff = models.BooleanField(default=False)
    fecha_creacion = models.DateTimeField(auto_now_add=True)

    objects = UsuarioManager()

    USERNAME_FIELD = 'correo'
    REQUIRED_FIELDS = ['nombre', 'rol']

    class Meta:
        db_table = 'usuarios'
        verbose_name = 'Usuario'
        verbose_name_plural = 'Usuarios'

    def __str__(self):
        return f"{self.nombre} - {self.rol}"

class Membresia(models.Model):
    nombre = models.CharField(max_length=100)
    duracion_dias = models.IntegerField()
    precio = models.DecimalField(max_digits=10, decimal_places=2)
    descripcion = models.TextField(blank=True, null=True)
    activa = models.BooleanField(default=True)
    fecha_creacion = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = 'membresias'
        verbose_name = 'Membresía'
        verbose_name_plural = 'Membresías'

    def __str__(self):
        return f"{self.nombre} - ${self.precio} COP"

class Cliente(models.Model):
    TIPO_DOCUMENTO_CHOICES = [
        ('CC', 'Cédula de Ciudadanía'),
        ('CE', 'Cédula de Extranjería'),
        ('TI', 'Tarjeta de Identidad'),
    ]
    
    ESTADOS = [
        ('activo', 'Activo'),
        ('inactivo', 'Inactivo'),
        ('pendiente', 'Pendiente'),
    ]

    # Tipo y número de documento
    tipo_documento = models.CharField(max_length=2, choices=TIPO_DOCUMENTO_CHOICES, default='CC')
    documento = models.CharField(max_length=20, primary_key=True)
    
    # Información personal
    nombres = models.CharField(max_length=100)
    apellidos = models.CharField(max_length=100)
    peso = models.DecimalField(max_digits=5, decimal_places=2, null=True, blank=True)
    fecha_nacimiento = models.DateField(null=True, blank=True)
    email = models.EmailField(null=True, blank=True)
    celular = models.CharField(max_length=15, null=True, blank=True)
    
    # Membresía
    membresia_actual = models.ForeignKey('Membresia', on_delete=models.SET_NULL, null=True, blank=True)
    fecha_inicio_membresia = models.DateField(null=True, blank=True)
    fecha_fin_membresia = models.DateField(null=True, blank=True)
    
    # Estado y registro
    estado = models.CharField(max_length=10, choices=ESTADOS, default='pendiente')
    fecha_registro = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = 'clientes'
        verbose_name = 'Cliente'
        verbose_name_plural = 'Clientes'
        ordering = ['-fecha_registro']

    def __str__(self):
        return f"{self.nombres} {self.apellidos} - {self.get_tipo_documento_display()}: {self.documento}"
    
    def get_tipo_documento_display_short(self):
        """Retorna solo las siglas del tipo de documento"""
        return self.tipo_documento

    def dias_para_vencer(self):
        """Retorna los días que faltan para que venza la membresía"""
        if self.fecha_fin_membresia:
            delta = self.fecha_fin_membresia - timezone.now().date()
            return delta.days
        return None
    
    def esta_por_vencer(self, dias=7):
        """Verifica si la membresía está por vencer en X días"""
        dias_restantes = self.dias_para_vencer()
        if dias_restantes is not None:
            return 0 <= dias_restantes <= dias
        return False
    
    def membresia_vencida(self):
        """Verifica si la membresía está vencida"""
        if self.fecha_fin_membresia:
            return self.fecha_fin_membresia < timezone.now().date()
        return True

    def renovar_membresia(self, membresia):
        """Renueva la membresía del cliente"""
        self.membresia_actual = membresia
        self.fecha_inicio_membresia = timezone.now().date()
        self.fecha_fin_membresia = timezone.now().date() + timedelta(days=membresia.duracion_dias)
        self.estado = 'activo'
        self.save()

class HistorialMembresia(models.Model):
    cliente = models.ForeignKey(Cliente, on_delete=models.CASCADE, related_name='historial')
    membresia = models.ForeignKey(Membresia, on_delete=models.SET_NULL, null=True)
    fecha_inicio = models.DateField()
    fecha_fin = models.DateField()
    precio_pagado = models.DecimalField(max_digits=10, decimal_places=2)
    fecha_registro = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = 'historial_membresias'
        verbose_name = 'Historial de Membresía'
        verbose_name_plural = 'Historial de Membresías'
        ordering = ['-fecha_registro']

    def __str__(self):
        return f"{self.cliente} - {self.membresia}"

class Asistencia(models.Model):
    cliente = models.ForeignKey(Cliente, on_delete=models.CASCADE)
    fecha = models.DateField(auto_now_add=True)
    hora = models.TimeField(auto_now_add=True)
    usuario_registro = models.ForeignKey(Usuario, on_delete=models.SET_NULL, null=True)

    class Meta:
        db_table = 'asistencias'
        verbose_name = 'Asistencia'
        verbose_name_plural = 'Asistencias'
        ordering = ['-fecha', '-hora']

    def __str__(self):
        return f"{self.cliente} - {self.fecha} {self.hora}"

class Pago(models.Model):
    METODOS_PAGO = [
        ('efectivo', 'Efectivo'),
        ('tarjeta', 'Tarjeta de Crédito/Débito'),
        ('transferencia', 'Transferencia Bancaria'),
        ('nequi', 'Nequi'),
        ('daviplata', 'Daviplata'),
    ]
    
    ESTADOS_PAGO = [
        ('pendiente', 'Pendiente'),
        ('validado', 'Validado'),
        ('rechazado', 'Rechazado'),
        ('cancelado', 'Cancelado'),
    ]
    
    TIPOS_PAGO = [
        ('membresia', 'Membresía'),
        ('renovacion', 'Renovación'),
    ]
    
    cliente = models.ForeignKey(Cliente, on_delete=models.CASCADE, related_name='pagos')
    membresia = models.ForeignKey(Membresia, on_delete=models.SET_NULL, null=True, blank=True)
    concepto = models.CharField(max_length=200)
    tipo_pago = models.CharField(max_length=20, choices=TIPOS_PAGO, default='membresia')
    monto = models.DecimalField(max_digits=10, decimal_places=2)
    metodo_pago = models.CharField(max_length=20, choices=METODOS_PAGO)
    estado = models.CharField(max_length=20, choices=ESTADOS_PAGO, default='pendiente')
    comprobante = models.CharField(max_length=100, blank=True, null=True)
    observaciones = models.TextField(blank=True, null=True)
    fecha_pago = models.DateTimeField(auto_now_add=True)
    fecha_validacion = models.DateTimeField(null=True, blank=True)
    usuario_registro = models.ForeignKey(Usuario, on_delete=models.SET_NULL, null=True, related_name='pagos_registrados')
    usuario_validacion = models.ForeignKey(Usuario, on_delete=models.SET_NULL, null=True, blank=True, related_name='pagos_validados')

    class Meta:
        db_table = 'pagos'
        verbose_name = 'Pago'
        verbose_name_plural = 'Pagos'
        ordering = ['-fecha_pago']

    def __str__(self):
        return f"Pago #{self.id} - {self.cliente} - ${self.monto}"

    def validar_pago(self, usuario_validacion):
        """Validar el pago y activar al cliente"""
        self.estado = 'validado'
        self.fecha_validacion = timezone.now()
        self.usuario_validacion = usuario_validacion
        self.save()
        
        # CAMBIAR ESTADO DEL CLIENTE A ACTIVO
        if self.cliente:
            self.cliente.estado = 'activo'
            self.cliente.save()
    
    def rechazar_pago(self, usuario_validacion, observacion):
        """Rechazar el pago y marcar cliente como inactivo"""
        self.estado = 'rechazado'
        self.fecha_validacion = timezone.now()
        self.usuario_validacion = usuario_validacion
        self.observaciones = observacion
        self.save()
        
        # CAMBIAR ESTADO DEL CLIENTE A INACTIVO
        if self.cliente:
            self.cliente.estado = 'inactivo'
            self.cliente.save()


class Bono(models.Model):
    """Modelo para gestionar bonos/regalos de días adicionales (máximo 3 días)"""
    
    TIPOS_BONO = [
        ('1_dia', '1 Día Extra'),
        ('2_dias', '2 Días Extra'),
        ('3_dias', '3 Días Extra'),
    ]
    
    cliente = models.ForeignKey(Cliente, on_delete=models.CASCADE, related_name='bonos')
    tipo_bono = models.CharField(max_length=20, choices=TIPOS_BONO)
    dias_regalo = models.IntegerField(default=0)
    motivo = models.CharField(max_length=200)
    fecha_otorgado = models.DateTimeField(auto_now_add=True)
    usuario_otorgo = models.ForeignKey(Usuario, on_delete=models.SET_NULL, null=True, blank=True)
    aplicado = models.BooleanField(default=False)
    fecha_aplicado = models.DateTimeField(null=True, blank=True)
    
    class Meta:
        db_table = 'bonos'
        ordering = ['-fecha_otorgado']
    
    def __str__(self):
        return f"Bono {self.get_tipo_bono_display()} - {self.cliente.nombres}"
    
    def aplicar_bono(self):
        """Aplica el bono sumando días a la membresía del cliente"""
        if not self.aplicado and self.cliente.fecha_fin_membresia:
            from datetime import timedelta
            from django.utils import timezone
            
            self.cliente.fecha_fin_membresia += timedelta(days=self.dias_regalo)
            self.cliente.save()
            
            self.aplicado = True
            self.fecha_aplicado = timezone.now()
            self.save()
            
            return True
        return False