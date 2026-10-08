# RetroTalk 2000

Chat interno de demostración construido con Google Apps Script y Google Sheets.

## Instalación
1. Crea una hoja de cálculo de Google vacía.
2. Abre Extensiones > Apps Script.
3. Crea `Code.gs`, `Index.html` y `appsscript.json` con el contenido de este paquete.
4. En Configuración del proyecto, activa la visualización del manifiesto si necesitas editar `appsscript.json`.
5. Ejecuta una vez `setupSheets_` desde el editor y autoriza el proyecto.
6. Selecciona Implementar > Nueva implementación > Aplicación web.
7. Ejecutar como: usuario que accede. Acceso: usuarios de tu dominio.
8. Abre la URL `/exec` con dos usuarios del mismo dominio para probar.

## Notificaciones
La app usa tres capas: aviso dentro de la interfaz, cambio temporal del título y Notification API del navegador. El usuario debe seleccionar “Activar avisos”. Algunos navegadores bloquean notificaciones dentro de iframes; el aviso interno seguirá funcionando.

## Arquitectura y límites
- Base de datos: pestañas `Mensajes` y `Usuarios`.
- Actualización: sondeo cada 3 segundos. Apps Script no ofrece un servidor WebSocket persistente.
- Concurrencia: `LockService` protege escrituras simultáneas.
- Seguridad: despliega solo para el dominio y no publiques como acceso anónimo.
- Escala: apropiado para una demo o equipo pequeño. Para mucho tráfico usa Firebase/Firestore o un backend en tiempo real.
