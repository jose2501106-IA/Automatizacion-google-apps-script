# Automatización de Procesos con Google Apps Script

Este proyecto es una demostración práctica de cómo aplicar principios de **DataOps** para eliminar el trabajo manual, asegurar la integridad de los datos y optimizar un flujo de trabajo completo a través de la automatización.

## 🚀 Problema Resuelto

El proceso manual de registrar asistentes para un evento, crear pases personalizados con códigos QR y enviar confirmaciones por correo era lento, propenso a errores humanos y consumía horas de trabajo valioso.

## 💡 Solución Implementada

Desarrollé un sistema 100% autónomo utilizando Google Apps Script que gestiona el ciclo de vida completo de los datos de registro de cada asistente. El flujo de trabajo es el siguiente:

1.  **Lectura de Datos:** Se activa automáticamente al recibir un nuevo registro en un Google Sheet.
2.  **Generación de Pase:** Crea una presentación personalizada en Google Slides usando una plantilla.
3.  **Creación de QR:** Se comunica con una API externa para generar un código QR único con la información del asistente.
4.  **Conversión a PDF:** Transforma la presentación en un archivo PDF listo para ser enviado.
5.  **Notificación por Email:** Envía un correo electrónico personalizado al asistente con su pase en PDF adjunto.
6.  **Actualización de Estado:** Modifica la fila original en el Google Sheet para marcar el registro como "Procesado", garantizando la trazabilidad.

## 🛠️ Tecnologías Utilizadas
* **Lenguaje:** JavaScript (Google Apps Script)
* **Plataforma:** Google Workspace (Sheets, Slides, Drive, Gmail)
* **APIs:** API externa para la generación de códigos QR (api.qrserver.com)

## 🏆 Impacto y Logros Clave
-   **Logro:** Eliminación del **100% del trabajo manual** y los errores asociados.
-   **Eficiencia:** Desarrollo y depuración de código acelerados mediante el uso estratégico de **Prompt Engineering**.
-   **Trazabilidad:** Creación de un sistema auditable donde el estado de cada registro es visible en tiempo real.
