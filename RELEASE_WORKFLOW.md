# 📝 **Checklist para generar una nueva versión (release)**

Este es el procedimiento oficial para crear una nueva versión del proyecto **extract_camiones** usando el flujo automatizado de versiones y empaquetado.

---

## ✅ **1. Actualizar el número de versión**

Ejecutar el script que incrementa automáticamente el último dígito:

```bash
python tools/bump_version.py
```

Esto convierte, por ejemplo:

```
0.4.2  →  0.4.3
```

La versión se actualiza en:

* `pyproject.toml`
* `src/version.py` la leerá automáticamente
* La GUI mostrará la nueva versión al ejecutarla

---

## ✅ **2. Generar el archivo ZIP listo para distribuir**

Ejecutar el generador de releases:

```bash
python -m tools.make_release
```

Esto creará en `dist/` un archivo con nombre automático:

```
extract_camiones_vX.Y.Z_YYYYMMDD_HHMMSS.zip
```

Este archivo **NO** se commitea al repositorio.

---

## ✅ **3. Guardar los cambios en Git**

Verificar qué archivos cambiaron:

```bash
git status
```

Lo esperado es ver solo:

```
modified: pyproject.toml
```

Agregar y commitear:

```bash
git add pyproject.toml
 to X.Y.Z"
```

Subir a GitHub:
Ejemplo
```bash
git pushgit commit -m "Bump version
```

---
git add RELEASE_WORKFLOW.md
git commit -m "Aclara pasos de commit en el workflow de release"
git push

## ✅ **4. Crear el tag correspondiente a la versión**

```bash
git tag vX.Y.Z
git push --tags
```

Esto permite que GitHub reconozca la versión formalmente.

---

## ✅ **5. Publicar el Release en GitHub**

1. Ir al repositorio → pestaña **Releases**
2. Clic en **"Draft a new release"**
3. Elegir el tag `vX.Y.Z` (si no existe, crearlo ahí mismo)
4. Título sugerido:

   ```
   EXTRACT_CAMIONES vX.Y.Z (Beta)
   ```
5. En la sección **Assets**, arrastrar el ZIP generado en `dist/`
6. Publicar con **Publish release**

