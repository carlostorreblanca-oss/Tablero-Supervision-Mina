# Tablero de Supervisión (COMINVI · Mina Aranzazu)

Acceso solo con cuenta de Microsoft 365. Esta página **no contiene datos**: al iniciar sesión, lee el personal desde OneDrive
(carpeta "Tablero Supervisores - datos") con los permisos de la persona. Sin permiso, no ve nada.

- `acceso.js`: inicio de sesión (MSAL), cierre tras 10 minutos sin uso y panel "Administrar accesos".
- `msal-browser.min.js`: biblioteca oficial de Microsoft (@azure/msal-browser 4.30.0).
- La configuración (clientId, ids de OneDrive) está en `<script id="conf">` de `index.html`.
