# ⏰ Complemento Excel Web - Insertar Hora Actual

Un complemento moderno para Excel Online (Office 365) que inserta la hora actual en celdas con múltiples formatos.

## 🎯 Inicio Rápido

### 1️⃣ Preparar los archivos

```
/archivos/
├── manifest.xml          ← Configuración del complemento
├── taskpane.html         ← Interfaz (HTML)
├── taskpane.css          ← Estilos (CSS)
└── taskpane.js           ← Lógica (JavaScript)
```

### 2️⃣ Alojar en servidor HTTPS

**Opción A: Usa GitHub Pages (Gratuito)**
1. Crea repositorio en GitHub
2. Copia los archivos
3. Habilita GitHub Pages
4. Tu URL: `https://tuuser.github.io/repo/taskpane.html`

**Opción B: Usa tu servidor web**
- Sube los archivos con FTP/SFTP
- URL debe ser HTTPS
- Actualiza `manifest.xml` con tu URL

### 3️⃣ Instalar en Excel

1. Abre **Excel Online**
2. **Insertar** → **Complementos** → **Mis complementos cargados**
3. **+ Cargar complemento personalizado**
4. Pega: `https://tu-url/manifest.xml`
5. ¡Listo! 🎉

---

## ✨ Características

- ⏰ Muestra hora en tiempo real
- 🎨 4 formatos diferentes
  - `hh:mm:ss`
  - `hh:mm AM/PM`
  - `mm:ss`
  - `dd/mm/yyyy hh:mm:ss`
- 📱 Interfaz responsive
- ✅ Confirmación visual
- 🔒 Seguro y privado

---

## 🔧 Desarrollo Local

### Con Python:
```bash
cd /ruta/archivos
python -m http.server 8000
```

### Con Node.js:
```bash
npm install
npm start
```

**Nota:** Excel Web requiere HTTPS. Para desarrollo local, usa **ngrok**:
```bash
ngrok http 8000
```

---

## 📝 Personalización

### Cambiar colores en `taskpane.css`:

```css
.btn-primary {
    background-color: #0078d4;  /* Cambiar aquí */
}

.info-section {
    background: linear-gradient(135deg, #0078d4 0%, #107c10 100%);
}
```

### Agregar más formatos en `taskpane.js`:

```javascript
case 'miFormato':
    valor = formatearMiFormato(ahora);
    break;
```

---

## 🐛 Problemas Comunes

| Problema | Solución |
|----------|----------|
| "No se carga" | Verifica URL en manifest.xml (¿es HTTPS?) |
| "No aparece el complemento" | Recarga Excel (Ctrl+F5) |
| "Error de permisos" | El servidor debe permitir CORS |
| "La celda no se actualiza" | Verifica permisos del archivo |

---

## 📚 Más Información

Lee **GUIA_INSTALACION_EXCEL_WEB.md** para:
- Instrucciones detalladas paso a paso
- Opciones avanzadas (Azure, SharePoint)
- Solución de problemas
- Publicación en Office Store

---

## 📋 Contenido de archivos

### manifest.xml
Archivo de configuración que Excel usa para:
- Identificar el complemento
- Cargar la interfaz
- Definir permisos

### taskpane.html
Panel lateral del complemento con:
- Pantalla de hora
- Botones de acción
- Opciones de formato
- Mensajes de estado

### taskpane.css
Estilos profesionales:
- Colores de Microsoft Design
- Animaciones suaves
- Diseño responsive
- Temas oscuro/claro compatible

### taskpane.js
Lógica con Office JavaScript API:
- `Office.onReady()` - Inicialización
- `Excel.run()` - Acceso a celdas
- Funciones de formato
- Manejo de errores

---

## 🚀 Próximas Mejoras

- [ ] Agregar botón en la cinta
- [ ] Guardar preferencias del usuario
- [ ] Insertar en múltiples celdas
- [ ] Sincronizar con zona horaria
- [ ] Traducción a más idiomas

---

## 📄 Licencia

MIT - Úsalo libremente

---

## 💡 Tips

✅ Usa HTTPS siempre (obligatorio en Excel Web)
✅ Comprueba CORS en tu servidor
✅ Usa herramientas de desarrollador (F12) para debug
✅ Lee comentarios en el código
✅ Visita docs.microsoft.com para más ejemplos

---

## 🤝 Soporte

Para más ayuda:
- [Documentación Office Add-ins](https://docs.microsoft.com/office/dev/add-ins/)
- [Excel JavaScript API](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
- Issues/Preguntas en GitHub

---

**¿Listo para empezar?** 🚀

👉 Lee la **GUIA_INSTALACION_EXCEL_WEB.md** para instrucciones completas
