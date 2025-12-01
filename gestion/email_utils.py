from django.core.mail import send_mail, send_mass_mail
from django.template.loader import render_to_string
from django.utils.html import strip_tags
from django.conf import settings
from .models import Cliente
from django.utils import timezone
from datetime import timedelta


class EmailService:
    """Servicio para envío de correos electrónicos"""
    
    @staticmethod
    def enviar_email_renovacion(cliente):
        """Enviar email de confirmación de renovación de membresía"""
        try:
            subject = f'¡Gracias por renovar tu membresía en FITTECH!'
            
            html_message = render_to_string('emails/renovacion.html', {
                'cliente': cliente,
                'membresia': cliente.membresia_actual,
                'fecha_inicio': cliente.fecha_inicio_membresia,
                'fecha_fin': cliente.fecha_fin_membresia,
            })
            
            plain_message = strip_tags(html_message)
            from_email = settings.DEFAULT_FROM_EMAIL
            to_email = cliente.email
            
            send_mail(
                subject,
                plain_message,
                from_email,
                [to_email],
                html_message=html_message,
                fail_silently=False,
            )
            return True
        except Exception as e:
            print(f"Error al enviar email de renovación: {str(e)}")
            return False
    
    @staticmethod
    def enviar_email_vencimiento(cliente):
        """Enviar email de recordatorio de vencimiento de membresía"""
        try:
            dias_restantes = (cliente.fecha_fin_membresia - timezone.now().date()).days
            
            subject = f'¡Tu membresía en FITTECH vence en {dias_restantes} días!'
            
            html_message = render_to_string('emails/vencimiento.html', {
                'cliente': cliente,
                'dias_restantes': dias_restantes,
                'fecha_vencimiento': cliente.fecha_fin_membresia,
                'membresia': cliente.membresia_actual,
            })
            
            plain_message = strip_tags(html_message)
            from_email = settings.DEFAULT_FROM_EMAIL
            to_email = cliente.email
            
            send_mail(
                subject,
                plain_message,
                from_email,
                [to_email],
                html_message=html_message,
                fail_silently=False,
            )
            return True
        except Exception as e:
            print(f"Error al enviar email de vencimiento: {str(e)}")
            return False
    
    @staticmethod
    def enviar_emails_masivos_vencimiento():
        """Enviar emails masivos a clientes con membresía por vencer"""
        dias_aviso = getattr(settings, 'DIAS_AVISO_VENCIMIENTO', 7)
        fecha_limite = timezone.now().date() + timedelta(days=dias_aviso)
        
        clientes_por_vencer = Cliente.objects.filter(
            fecha_fin_membresia__lte=fecha_limite,
            fecha_fin_membresia__gte=timezone.now().date(),
            estado='activo'
        )
        
        emails_enviados = 0
        emails_fallidos = 0
        
        messages = []
        for cliente in clientes_por_vencer:
            try:
                dias_restantes = (cliente.fecha_fin_membresia - timezone.now().date()).days
                
                subject = f'¡Tu membresía en FITTECH vence en {dias_restantes} días!'
                
                html_message = render_to_string('emails/vencimiento.html', {
                    'cliente': cliente,
                    'dias_restantes': dias_restantes,
                    'fecha_vencimiento': cliente.fecha_fin_membresia,
                    'membresia': cliente.membresia_actual,
                })
                
                plain_message = strip_tags(html_message)
                from_email = settings.DEFAULT_FROM_EMAIL
                
                messages.append((subject, plain_message, from_email, [cliente.email]))
                emails_enviados += 1
            except Exception as e:
                print(f"Error preparando email para {cliente.email}: {str(e)}")
                emails_fallidos += 1
        
        if messages:
            try:
                send_mass_mail(messages, fail_silently=False)
            except Exception as e:
                print(f"Error en envío masivo: {str(e)}")
                return {'enviados': 0, 'fallidos': len(messages), 'total_clientes': 0}
        
        return {
            'enviados': emails_enviados,
            'fallidos': emails_fallidos,
            'total_clientes': clientes_por_vencer.count()
        }
    
    @staticmethod
    def enviar_email_bienvenida(cliente):
        """Enviar email de bienvenida a nuevo cliente"""
        try:
            subject = f'¡Bienvenido a FITTECH, {cliente.nombres}!'
            
            html_message = render_to_string('emails/bienvenida.html', {
                'cliente': cliente,
                'membresia': cliente.membresia_actual,
                'fecha_inicio': cliente.fecha_inicio_membresia,
                'fecha_fin': cliente.fecha_fin_membresia,
            })
            
            plain_message = strip_tags(html_message)
            from_email = settings.DEFAULT_FROM_EMAIL
            to_email = cliente.email
            
            send_mail(
                subject,
                plain_message,
                from_email,
                [to_email],
                html_message=html_message,
                fail_silently=False,
            )
            return True
        except Exception as e:
            print(f"Error al enviar email de bienvenida: {str(e)}")
            return False

    @staticmethod
    def enviar_email_reactivacion(cliente):
        """Enviar email de reactivación a cliente inactivo, con HTML"""
        try:
            subject = '¡Te extrañamos en FITTECH! - Vuelve a entrenar con nosotros'
            hoy = timezone.now().date()
            html_message = render_to_string('emails/inactivo.html', {
                'cliente': cliente,
                'hoy': hoy,
                'settings': settings,
            })
            plain_message = strip_tags(html_message)
            from_email = settings.DEFAULT_FROM_EMAIL
            to_email = cliente.email

            send_mail(
                subject,
                plain_message,
                from_email,
                [to_email],
                html_message=html_message,
                fail_silently=False,
            )
            print(f"✓ Email enviado exitosamente a {cliente.email}")
            return True
        except Exception as e:
            print(f"✗ ERROR enviando email a {cliente.email}: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
