# Port Python del instalador de plantillas de Office

Esta carpeta contiene una versión en Python del flujo definido por los scripts
BAT originales en `Script/`. El código reproduce la resolución de rutas,
validación de autores y orquestación de copias/MRU, manteniendo intactos los
`.bat` originales.

## Mapeo de entrypoints

| Script BAT                                     | Módulo Python y puntos clave                               |
| ---------------------------------------------- | ---------------------------------------------------------- |
| `1-2. MainInstaller.bat`                       | `main_installer.py` (`main`, `resolve_base_directory`)     |
| `1-2. AuthContainerTools.bat`                  | `auth_tools.py` (`check_template_author`)                  |
| `1-2. MRU-PathResolver.bat`                    | `mru_resolver.py` (`detect_mru_path`, `record_recent_template`) |
| `1-2. Repair Office template MRU.bat`          | `mru_repair.py` (`repair_template_mru`)                    |
| `1-2. TemplateFolderOpener.bat`                | `folder_opener.py` (`open_template_folders`)               |
| `1-2. ResolveAppProperties.bat` (rutas Office) | `office_paths.py` (`detect_office_paths`)                  |

## Flujo principal

1. **Resolución de la pista/base** (`resolve_base_directory`):
   - Usa el argumento CLI o `PIN_LAUNCHER_DIR`; si faltan, toma la ruta del script.
   - Busca plantillas directamente o en subcarpetas `payload/`, `templates/` o
     `extracted/` para fijar el directorio base.
   - Si se ejecuta desde `%APPDATA%` sin pista explícita se muestra el error en
     español original para evitar rutas vacías.
2. **Validación y entorno**:
   - Se configuran los autores permitidos desde `AllowedTemplateAuthors` o la
     lista por defecto (`www.grada.cc;www.gradaz.com`).
   - Se verifica el entorno (`check_environment`) y se cierran aplicaciones de
     Office si se ejecuta en Windows.
   - Se limpia la MRU de plantillas (`repair_template_mru`).
3. **Instalación de plantillas base** (`install_base_templates`):
   - Copia `Normal*`, `Blank*`, `Book*` y `Sheet*` hacia las rutas estándar
     (`%APPDATA%\Microsoft\Templates` y `%APPDATA%\Microsoft\Excel\XLSTART`).
   - Se crean respaldos en `Backup/` antes de sobreescribir y se marcan las
     carpetas a abrir según corresponda.
4. **Detección de rutas personalizadas** (`detect_office_paths`):
   - Consulta el registro (cuando hay soporte) para `PersonalTemplates` o
     `UserTemplates`; si no existen, crea `Documentos\Custom Templates` como
     predeterminado y prepara la carpeta de Document Themes.
5. **Copia de plantillas personalizadas** (`copy_custom_templates`):
   - Recorre todos los `.dotm/.dotx/.potm/.potx/.xltm/.xltx/.thmx` que no sean
     plantillas base, valida autores y los copia a Word/PowerPoint/Excel o a
     la carpeta de temas.
   - Actualiza flags para abrir carpetas (custom, roaming, Excel startup) y
     solicita registrar MRU de forma segura.
6. **Apertura de carpetas y apps**:
   - Llama a `open_template_folders` con los flags calculados; en Windows usa
     Explorer y en otros SO deja trazas de advertencia.
   - Aplica el retardo `DOCUMENT_THEME_OPEN_DELAY_SECONDS` antes de relanzar
     Word/PowerPoint/Excel cuando procede.

## Uso de la CLI

```bash
python -m python_port.main_installer "C:\ruta\a\plantillas" \
  --design-mode \
  --allowed-authors "autor1;autor2" \
  --document-theme-delay 1
```

Opciones útiles:

- `--check-author <ruta>`: sólo valida autor de archivo o carpeta y termina
  devolviendo `TRUE/FALSE` como en el BAT.
- `--disable-author-validation`: replica `AuthorValidationEnabled=FALSE`.
- `--design-mode`: emite mensajes detallados como `IsDesignModeEnabled=true`.

Variables de entorno reconocidas: `PIN_LAUNCHER_DIR`, `AllowedTemplateAuthors`,
`AuthorValidationEnabled`, `DOCUMENT_THEME_OPEN_DELAY_SECONDS` e
`IsDesignModeEnabled`. En Windows también se respetan `%APPDATA%` y `%USERPROFILE%`
para calcular rutas por defecto.

## Dependencias

- Python 3.11+.
- Windows para interacción real con el registro, MRU y apertura de carpetas.
  En otros sistemas las operaciones sensibles se simulan y se advierte por
  consola en modo diseño.
