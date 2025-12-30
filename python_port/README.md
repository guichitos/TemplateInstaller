# Port Python simplificado

Tres archivos reproducen las tareas de instalación/desinstalación sin depender
de la estructura del BAT original:

- `installer.py`: instala plantillas tomando como base la carpeta desde la que
  se ejecuta (busca plantillas también en `payload/`, `templates/` o
  `extracted/`). Replica validación de autores, copia plantillas base y
  personalizadas, aplica respaldos y abre apps según flags.
- `uninstaller.py`: elimina las plantillas base y las copias personalizadas
  detectadas, usando la misma carpeta actual como referencia.
- `common.py`: constantes, validación de autores, utilidades de copia,
  respaldo, cierres de Office y arranque opcional de apps.

## Uso rápido

Ejecutar desde la carpeta que contiene las plantillas:

```bash
python -m python_port.installer --design-mode
python -m python_port.uninstaller --design-mode
```

Opciones clave:

- `--check-author <ruta>` (solo instalador): muestra `TRUE/FALSE` y termina.
- `--allowed-authors "autor1;autor2"`: sustituye la lista por defecto.
- `--design-mode`: salida detallada en consola.

## Variables de entorno soportadas

- `AllowedTemplateAuthors`: lista separada por `;`.
- `AuthorValidationEnabled`: `TRUE` (default) / `FALSE`.
- `IsDesignModeEnabled`: `true/false` (predetermina el modo diseño).
- `DOCUMENT_THEME_OPEN_DELAY_SECONDS`: retraso antes de abrir apps.
- Rutas opcionales: `CUSTOM_OFFICE_TEMPLATE_PATH`,
  `CUSTOM_OFFICE_ADDITIONAL_TEMPLATE_PATH`, `ROAMING_TEMPLATE_FOLDER_PATH`,
  `EXCEL_STARTUP_FOLDER_PATH`.
