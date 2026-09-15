# Papelería ITM V7.1 - Versión segura

Sistema Streamlit para inventario, solicitudes, salidas, compras, áreas, usuarios, kardex y reportes.

## Seguridad
La contraseña del administrador ya NO está escrita en `app.py` ni se muestra en la pantalla de acceso.

En Streamlit Cloud configure en **Settings → Secrets**:

```toml
ADMIN_USER = "admin"
ADMIN_PASSWORD = "SU_CLAVE_SEGURA"
```

No suba `.streamlit/secrets.toml` a GitHub.

## Importante sobre la base de datos
La base SQLite se crea localmente y está excluida de Git mediante `.gitignore`. Para una operación institucional permanente en Streamlit Cloud se recomienda migrar a una base de datos externa/persistente.

## Ejecutar localmente
```powershell
cd "C:\Users\juang\OneDrive\Desktop\papeleria_ITM_v7_1_SEGURA"
python -m streamlit run app.py
```

## Excel
El archivo `papeleria listado.xlsx` debe permanecer en la raíz del repositorio junto a `app.py`.


## Credenciales iniciales V7.2

Configura estas credenciales en **Streamlit Cloud → Manage app → Settings → Secrets**. No las publiques en GitHub.

### Administrador
- Usuario: `admin`
- Clave: `ITM_Admin_2026!`
- Rol: Administrador

### Usuario de consulta
- Usuario: `consulta`
- Clave: `ITM_Consulta_2026!`
- Rol: Consulta

El usuario `consulta` puede consultar inventario, valores/IVA, realizar pedidos, historial y reportes, pero no puede registrar compras ni entrar en Administración.
