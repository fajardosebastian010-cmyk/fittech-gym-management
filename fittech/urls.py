from django.contrib import admin
from django.urls import path
from gestion import controllers


urlpatterns = [
    # Admin de Django
    path('admin/', admin.site.urls),
    
    # ============= LOGIN Y AUTENTICACIÓN =============
    path('', controllers.login_view, name='login'),
    path('login/', controllers.login_view, name='login'),
    path('logout/', controllers.logout_view, name='logout'),
    
    # ============= DASHBOARD =============
    path('dashboard/', controllers.dashboard, name='dashboard'),
    
    # ============= MEMBRESÍAS =============
    path('membresias/', controllers.membresias_listar, name='membresias_listar'),
    path('membresias/crear/', controllers.membresias_crear, name='membresias_crear'),
    path('membresias/<int:id>/', controllers.membresias_ver, name='membresias_ver'),
    path('membresias/<int:id>/editar/', controllers.membresias_editar, name='membresias_editar'),
    path('membresias/<int:id>/eliminar/', controllers.membresias_eliminar, name='membresias_eliminar'),
    
    # ============= CLIENTES =============
    path('clientes/', controllers.clientes_listar, name='clientes_listar'),
    path('clientes/crear/', controllers.clientes_crear, name='clientes_crear'),
    path('clientes/importar/', controllers.clientes_importar_excel, name='clientes_importar_excel'),
    path('clientes/<str:documento>/', controllers.clientes_ver, name='clientes_ver'),
    path('clientes/<str:documento>/editar/', controllers.clientes_editar, name='clientes_editar'),
    path('clientes/<str:documento>/eliminar/', controllers.clientes_eliminar, name='clientes_eliminar'),
    path('clientes/<str:documento>/renovar/', controllers.clientes_renovar, name='clientes_renovar'),
    path('clientes/<str:documento>/asistencias/', controllers.cliente_asistencias, name='cliente_asistencias'),
    
    # ============= ASISTENCIAS =============
    path('asistencias/', controllers.asistencias_listar, name='asistencias_listar'),
    path('asistencias/registrar/', controllers.asistencias_registrar, name='asistencias_registrar'),
    path('asistencias/exportar/excel/', controllers.asistencias_exportar_excel, name='asistencias_exportar_excel'),
    path('asistencias/exportar/pdf/', controllers.asistencias_exportar_pdf, name='asistencias_exportar_pdf'),
    
    # ============= PAGOS =============
    path('pagos/', controllers.pagos_listar, name='pagos_listar'),
    path('pagos/crear/', controllers.pagos_crear, name='pagos_crear'),
    path('pagos/registrar/<str:documento>/', controllers.pagos_registrar, name='pagos_registrar'),  # ✅ Solo una vez, no duplicar
    path('pagos/reportes/', controllers.pagos_reportes, name='pagos_reportes'),
    path('pagos/exportar/excel/', controllers.pagos_exportar_excel, name='pagos_exportar_excel'),
    path('pagos/exportar/pdf/', controllers.pagos_exportar_pdf, name='pagos_exportar_pdf'),
    path('pagos/<int:id>/', controllers.pagos_ver, name='pagos_ver'),
    path('pagos/<int:id>/editar/', controllers.pagos_editar, name='pagos_editar'),
    path('pagos/<int:id>/validar/', controllers.pagos_validar, name='pagos_validar'),
    path('pagos/<int:id>/eliminar/', controllers.pagos_eliminar, name='pagos_eliminar'),
    
    # ============= USUARIOS =============
    path('usuarios/', controllers.usuarios_listar, name='usuarios_listar'),
    path('usuarios/crear/', controllers.usuarios_crear, name='usuarios_crear'),
    path('usuarios/<int:id>/', controllers.usuarios_ver, name='usuarios_ver'),
    path('usuarios/<int:id>/editar/', controllers.usuarios_editar, name='usuarios_editar'),
    path('usuarios/<int:id>/eliminar/', controllers.usuarios_eliminar, name='usuarios_eliminar'),
    
    # ============= BONOS =============
    path('bonos/', controllers.bonos_listar, name='bonos_listar'),
    path('bonos/crear/', controllers.bonos_crear, name='bonos_crear'),
    path('bonos/estadisticas/', controllers.bonos_estadisticas, name='bonos_estadisticas'),
    path('bonos/<int:id>/aplicar/', controllers.bonos_aplicar, name='bonos_aplicar'),
    path('bonos/<int:id>/eliminar/', controllers.bonos_eliminar, name='bonos_eliminar'),
    
    # ============= REPORTES =============
    path('reportes/', controllers.reportes_generales, name='reportes_generales'),
    path('reportes/membresias/excel/', controllers.reportes_membresias_excel, name='reportes_membresias_excel'),
    path('reportes/membresias/pdf/', controllers.reportes_membresias_pdf, name='reportes_membresias_pdf'),
    path('reportes/clientes/excel/', controllers.reportes_clientes_excel, name='reportes_clientes_excel'),
    path('reportes/clientes/pdf/', controllers.reportes_clientes_pdf, name='reportes_clientes_pdf'),
    path('reportes/usuarios/excel/', controllers.reportes_usuarios_excel, name='reportes_usuarios_excel'),
    path('reportes/usuarios/pdf/', controllers.reportes_usuarios_pdf, name='reportes_usuarios_pdf'),
    path('reportes/consolidado/excel/', controllers.reporte_consolidado_excel, name='reporte_consolidado_excel'),
    
    # ============= EMAILS =============
    path('emails/panel/', controllers.emails_panel, name='emails_panel'),
    path('emails/enviar-vencimiento/', controllers.enviar_emails_vencimiento, name='enviar_emails_vencimiento'),
    path('emails/renovacion/<str:documento>/', controllers.enviar_email_renovacion_individual, name='enviar_email_renovacion_individual'),
    path('emails/vencimiento/<str:documento>/', controllers.enviar_email_vencimiento_individual, name='enviar_email_vencimiento_individual'),
    path('emails/clientes-inactivos/', controllers.emails_clientes_inactivos, name='emails_clientes_inactivos'),
    path('emails/enviar-inactivos/', controllers.enviar_emails_inactivos, name='enviar_emails_inactivos'),
    path('emails/reactivacion/<str:documento>/', controllers.enviar_email_reactivacion_individual, name='enviar_email_reactivacion_individual'),
]
