from django.utils import timezone
from django.db.models import Sum, Count, Avg
from datetime import timedelta, datetime
from .models import Usuario, Membresia, Cliente, Asistencia, Pago

class UsuarioDAO:
    """Data Access Object para gestionar Usuarios"""
    
    @staticmethod
    def obtener_todos():
        return Usuario.objects.all().order_by('-fecha_creacion')
    
    @staticmethod
    def obtener_por_id(id):
        return Usuario.objects.get(id=id)
    
    @staticmethod
    def obtener_por_correo(correo):
        return Usuario.objects.get(correo=correo)
    
    @staticmethod
    def crear(datos):
        usuario = Usuario.objects.create_user(
            correo=datos['correo'],
            password=datos['password'],
            nombre=datos['nombre'],
            rol=datos['rol']
        )
        return usuario
    
    @staticmethod
    def actualizar(id, datos):
        usuario = Usuario.objects.get(id=id)
        usuario.nombre = datos.get('nombre', usuario.nombre)
        usuario.correo = datos.get('correo', usuario.correo)
        usuario.rol = datos.get('rol', usuario.rol)
        
        if 'password' in datos and datos['password']:
            usuario.set_password(datos['password'])
        
        usuario.save()
        return usuario
    
    @staticmethod
    def eliminar(id):
        usuario = Usuario.objects.get(id=id)
        usuario.delete()
    
    @staticmethod
    def obtener_administradores():
        return Usuario.objects.filter(rol='administrador')
    
    @staticmethod
    def obtener_empleados():
        return Usuario.objects.filter(rol='empleado')


class MembresiaDAO:
    """Data Access Object para gestionar Membresías"""
    
    @staticmethod
    def obtener_todas():
        return Membresia.objects.filter(activa=True).order_by('-fecha_creacion')
    
    @staticmethod
    def obtener_por_id(id):
        return Membresia.objects.get(id=id)
    
    @staticmethod
    def crear(datos):
        membresia = Membresia.objects.create(**datos)
        return membresia
    
    @staticmethod
    def actualizar(id, datos):
        membresia = Membresia.objects.get(id=id)
        for key, value in datos.items():
            setattr(membresia, key, value)
        membresia.save()
        return membresia
    
    @staticmethod
    def eliminar(id):
        membresia = Membresia.objects.get(id=id)
        membresia.activa = False
        membresia.save()
    
    @staticmethod
    def obtener_activas():
        return Membresia.objects.filter(activa=True)
    
    @staticmethod
    def obtener_estadisticas():
        stats = {
            'total_membresias': Membresia.objects.filter(activa=True).count(),
            'precio_promedio': Membresia.objects.filter(activa=True).aggregate(Avg('precio'))['precio__avg'] or 0,
            'duracion_promedio': Membresia.objects.filter(activa=True).aggregate(Avg('duracion_dias'))['duracion_dias__avg'] or 0,
        }
        return stats


class ClienteDAO:
    """Data Access Object para gestionar Clientes"""
    
    @staticmethod
    def obtener_todos():
        return Cliente.objects.all().order_by('-fecha_registro')
    
    @staticmethod
    def obtener_por_documento(documento):
        return Cliente.objects.get(documento=documento)
    
    @staticmethod
    def crear(datos):
        """Crea un cliente con los datos proporcionados"""
        cliente = Cliente.objects.create(**datos)
        return cliente
    
    @staticmethod
    def actualizar(documento, datos):
        cliente = Cliente.objects.get(documento=documento)
        for key, value in datos.items():
            setattr(cliente, key, value)
        cliente.save()
        return cliente
    
    @staticmethod
    def eliminar(documento):
        cliente = Cliente.objects.get(documento=documento)
        cliente.delete()
    
    @staticmethod
    def obtener_activos():
        """Obtener clientes con estado activo"""
        return Cliente.objects.filter(estado='activo')
    
    @staticmethod
    def obtener_inactivos():
        """Obtener clientes con estado inactivo"""
        return Cliente.objects.filter(estado='inactivo')
    
    @staticmethod
    def obtener_pendientes():
        """Obtener clientes con estado pendiente (esperando validación de pago)"""
        return Cliente.objects.filter(estado='pendiente')
    
    @staticmethod
    def obtener_por_membresia(membresia):
        """Obtener clientes activos con una membresía específica"""
        return Cliente.objects.filter(membresia_actual=membresia, estado='activo')
    
    @staticmethod
    def obtener_clientes_por_vencer(dias=7):
        """Obtener clientes cuya membresía vence en los próximos días especificados"""
        hoy = timezone.now().date()
        fecha_limite = hoy + timedelta(days=dias)
        
        return Cliente.objects.filter(
            fecha_fin_membresia__lte=fecha_limite,
            fecha_fin_membresia__gte=hoy,
            estado='activo'
        ).order_by('fecha_fin_membresia')
    
    @staticmethod
    def obtener_clientes_vencidos():
        """Obtener clientes con membresía vencida"""
        hoy = timezone.now().date()
        
        return Cliente.objects.filter(
            fecha_fin_membresia__lt=hoy,
            estado='activo'
        ).order_by('fecha_fin_membresia')
    
    @staticmethod
    def obtener_estadisticas():
        """Obtener estadísticas generales de clientes"""
        hoy = timezone.now().date()
        
        stats = {
            'total_clientes': Cliente.objects.count(),
            'clientes_activos': Cliente.objects.filter(estado='activo').count(),
            'clientes_inactivos': Cliente.objects.filter(estado='inactivo').count(),
            'clientes_pendientes': Cliente.objects.filter(estado='pendiente').count(),
            'clientes_por_vencer': ClienteDAO.obtener_clientes_por_vencer().count(),
            'clientes_vencidos': ClienteDAO.obtener_clientes_vencidos().count(),
        }
        return stats
    
    @staticmethod
    def buscar(query):
        """Buscar clientes por nombre, apellido, documento o email"""
        from django.db.models import Q
        return Cliente.objects.filter(
            Q(documento__icontains=query) |
            Q(nombres__icontains=query) |
            Q(apellidos__icontains=query) |
            Q(email__icontains=query) |
            Q(celular__icontains=query)
        ).order_by('-fecha_registro')

class AsistenciaDAO:
    """Data Access Object para gestionar Asistencias"""
    
    @staticmethod
    def obtener_todas():
        return Asistencia.objects.all().order_by('-fecha', '-hora')
    
    @staticmethod
    def obtener_por_fecha(fecha):
        return Asistencia.objects.filter(fecha=fecha).order_by('-hora')
    
    @staticmethod
    def obtener_por_cliente(cliente):
        return Asistencia.objects.filter(cliente=cliente).order_by('-fecha', '-hora')
    
    @staticmethod
    def crear(cliente, usuario_registro):
        asistencia = Asistencia.objects.create(
            cliente=cliente,
            usuario_registro=usuario_registro
        )
        return asistencia
    
    @staticmethod
    def contar_asistencias_dia(fecha=None):
        """Contar asistencias de un día específico (por defecto hoy)"""
        if fecha is None:
            fecha = timezone.now().date()
        return Asistencia.objects.filter(fecha=fecha).count()
    
    @staticmethod
    def contar_asistencias_mes(fecha=None):
        """Contar asistencias del mes actual"""
        if fecha is None:
            fecha = timezone.now().date()
        
        inicio_mes = fecha.replace(day=1)
        return Asistencia.objects.filter(fecha__gte=inicio_mes, fecha__lte=fecha).count()
    
    @staticmethod
    def obtener_estadisticas():
        hoy = timezone.now().date()
        inicio_mes = hoy.replace(day=1)
        
        stats = {
            'total_asistencias': Asistencia.objects.count(),
            'asistencias_hoy': AsistenciaDAO.contar_asistencias_dia(),
            'asistencias_mes': Asistencia.objects.filter(fecha__gte=inicio_mes).count(),
        }
        return stats
    
    @staticmethod
    def obtener_reporte_rango(fecha_inicio, fecha_fin):
        """Obtener asistencias en un rango de fechas"""
        return Asistencia.objects.filter(
            fecha__gte=fecha_inicio,
            fecha__lte=fecha_fin
        ).order_by('-fecha', '-hora')


class PagoDAO:
    """Data Access Object para gestionar Pagos"""
    
    @staticmethod
    def obtener_todos():
        return Pago.objects.all().order_by('-fecha_pago')
    
    @staticmethod
    def obtener_pendientes():
        return Pago.objects.filter(estado='pendiente').order_by('-fecha_pago')
    
    @staticmethod
    def obtener_validados():
        return Pago.objects.filter(estado='validado').order_by('-fecha_pago')
    
    @staticmethod
    def obtener_rechazados():
        return Pago.objects.filter(estado='rechazado').order_by('-fecha_pago')
    
    @staticmethod
    def obtener_por_cliente(cliente):
        return Pago.objects.filter(cliente=cliente).order_by('-fecha_pago')
    
    @staticmethod
    def crear(datos):
        return Pago.objects.create(**datos)
    
    @staticmethod
    def actualizar(id, datos):
        pago = Pago.objects.get(id=id)
        for key, value in datos.items():
            setattr(pago, key, value)
        pago.save()
        return pago
    
    @staticmethod
    def eliminar(id):
        pago = Pago.objects.get(id=id)
        pago.delete()
    
    @staticmethod
    def obtener_estadisticas():
        """Obtener estadísticas de pagos - CORREGIDO CON DATETIME"""
        ahora = timezone.now()
        hoy = ahora.date()
        inicio_mes = hoy.replace(day=1)
        
        # Rangos de datetime
        inicio_dia = ahora.replace(hour=0, minute=0, second=0, microsecond=0)
        fin_dia = ahora.replace(hour=23, minute=59, second=59, microsecond=999999)
        inicio_mes_dt = ahora.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        
        # Solo contar pagos VALIDADOS con rangos de datetime
        stats = {
            'total_pagos': Pago.objects.filter(estado='validado').count(),
            'total_ingresos': float(Pago.objects.filter(estado='validado').aggregate(Sum('monto'))['monto__sum'] or 0),
            'pagos_pendientes': Pago.objects.filter(estado='pendiente').count(),
            'pagos_rechazados': Pago.objects.filter(estado='rechazado').count(),
            'pagos_hoy': Pago.objects.filter(
                fecha_pago__gte=inicio_dia,
                fecha_pago__lte=fin_dia,
                estado='validado'
            ).count(),
            'ingresos_hoy': float(Pago.objects.filter(
                fecha_pago__gte=inicio_dia,
                fecha_pago__lte=fin_dia,
                estado='validado'
            ).aggregate(Sum('monto'))['monto__sum'] or 0),
            'ingresos_mes': float(Pago.objects.filter(
                fecha_pago__gte=inicio_mes_dt,
                fecha_pago__lte=ahora,
                estado='validado'
            ).aggregate(Sum('monto'))['monto__sum'] or 0),
        }
        
        # Pagos por método (solo validados)
        pagos_por_metodo = Pago.objects.filter(estado='validado').values('metodo_pago').annotate(
            total=Count('id'),
            monto_total=Sum('monto')
        ).order_by('-monto_total')
        
        stats['pagos_por_metodo'] = list(pagos_por_metodo)
        
        # Pagos por tipo (solo validados)
        pagos_por_tipo = Pago.objects.filter(estado='validado').values('tipo_pago').annotate(
            total=Count('id'),
            monto_total=Sum('monto')
        ).order_by('-monto_total')
        
        stats['pagos_por_tipo'] = list(pagos_por_tipo)
        
        return stats
    
    @staticmethod
    def obtener_reporte_fechas(fecha_inicio, fecha_fin):
        """Obtener pagos validados en un rango de fechas"""
        inicio = datetime.strptime(fecha_inicio, '%Y-%m-%d').date()
        fin = datetime.strptime(fecha_fin, '%Y-%m-%d').date()
        
        return Pago.objects.filter(
            fecha_pago__date__gte=inicio,
            fecha_pago__date__lte=fin,
            estado='validado'
        ).order_by('-fecha_pago')
    
    @staticmethod
    def obtener_ingresos_por_mes(meses=6):
        """Obtener ingresos de los últimos N meses"""
        hoy = timezone.now().date()
        ingresos = []
        
        for i in range(meses - 1, -1, -1):
            fecha = hoy - timedelta(days=30 * i)
            inicio_mes = fecha.replace(day=1)
            
            if fecha.month == 12:
                fin_mes = fecha.replace(day=31)
            else:
                siguiente_mes = fecha.replace(day=28) + timedelta(days=4)
                fin_mes = siguiente_mes - timedelta(days=siguiente_mes.day)
            
            ingresos_mes = Pago.objects.filter(
                fecha_pago__date__gte=inicio_mes,
                fecha_pago__date__lte=fin_mes,
                estado='validado'
            ).aggregate(Sum('monto'))['monto__sum'] or 0
            
            ingresos.append({
                'mes': fecha.strftime('%B %Y'),
                'ingresos': float(ingresos_mes)
            })
        
        return ingresos
    
    @staticmethod
    def obtener_top_clientes(limit=10):
        """Obtener los clientes que más han pagado"""
        return Pago.objects.filter(estado='validado').values(
            'cliente__documento',
            'cliente__nombres',
            'cliente__apellidos'
        ).annotate(
            total_pagado=Sum('monto'),
            cantidad_pagos=Count('id')
        ).order_by('-total_pagado')[:limit]
