# 📊 Complemento para Insertar Hora en Excel Web

## ¿Qué es este complemento?

Este es un **complemento de Excel Web (Office 365 Add-in)** que permite insertar la hora actual en la celda activa de Excel Online. Está escrito en JavaScript usando la **Office JavaScript API**.

---

## 📋 Contenido de los Archivos

1. **manifest.xml** - Archivo de configuración del complemento
2. **taskpane.html** - Interfaz del complemento (panel lateral)
3. **taskpane.css** - Estilos visuales
4. **taskpane.js** - Lógica JavaScript con Office API

---

## 🚀 Opción 1: Alojar en Servidor Web (Recomendado)

### Paso 1: Preparar los archivos

```
tu-servidor/
├── manifest.xml
├── taskpane.html
├── taskpane.css
└── taskpane.js
```

### Paso 2: Actualizar la URL en manifest.xml

Abre `manifest.xml` y reemplaza:
```xml
<SourceLocation DefaultValue="https://tu-dominio.com/taskpane.html"/>
```

Por tu URL real:
```xml
<SourceLocation DefaultValue="https://tupagina.com/excel-addon/taskpane.html"/>
```

**Nota:** La URL debe ser **HTTPS** (no HTTP)

### Paso 3: Subir archivos al servidor

Sube los 4 archivos a tu servidor web usando:
- FTP, SFTP
- Gestor de archivos del hosting
- Git (si usas GitHub Pages o similar)

### Paso 4: Subir el manifesto a Excel

1. Abre **Excel Online** (office.com)
2. Crea o abre un libro
3. Ve a **Insertar** > **Complementos** > **Mis complementos** > **Mis complementos cargados**
4. Selecciona **+ Cargar un complemento personalizado**
5. Pega la URL completa del manifest:
   ```
   https://tu-servidor.com/manifest.xml
   ```
6. Presiona **Cargar** o **OK**

¡Listo! El complemento debería aparecer en el panel derecho.

---

## 🖥️ Opción 2: Usar Azure o SharePoint (Enterprise)

### Con Azure App Service:

1. **Crea una App Service** en Azure
2. **Despliega los archivos** (HTML, CSS, JS)
3. **Registra el complemento** en el Centro de administración de Microsoft 365
4. **Comparte la URL** del manifest con los usuarios

### Con SharePoint:

1. Sube los archivos a una **biblioteca de documentos**
2. Obtén la URL completa del manifest.xml
3. Sigue el proceso de carga en Excel Online

---

## 📱 Opción 3: Desarrollo Local (Para Pruebas)

Si solo necesitas probarlo localmente:

### Con Python (servidor simple):

```bash
cd /ruta/a/los/archivos
python -m http.server 8000
```

La URL será: `http://localhost:8000/taskpane.html`

**PROBLEMA:** Excel Web requiere HTTPS. Para desarrollo local:

### Con ngrok (tunel HTTPS):

1. **Descarga ngrok**: https://ngrok.com
2. **Ejecuta ngrok**:
   ```bash
   ngrok http 8000
   ```
3. **Usa la URL HTTPS** que ngrok genera (ej: `https://abc123.ngrok.io`)
4. **Actualiza manifest.xml** con esa URL

---

## ✨ Características del Complemento

### Funcionalidades:

✅ Inserta la hora actual en la celda activa
✅ Muestra la hora actualizada cada segundo
✅ Múltiples formatos de hora:
   - Hora estándar: hh:mm:ss
   - Hora 12h: hh:mm AM/PM
   - Solo minutos: mm:ss
   - Con fecha: dd/mm/yyyy hh:mm:ss

✅ Interfaz intuitiva y responsive
✅ Mensajes de confirmación
✅ Compatible con Excel Web (Office 365)

---

## 💻 Cómo Usar el Complemento

### Una vez instalado:

1. **Abre Excel Online**
2. **Haz clic en Insertar** > **Complementos** > Tu complemento
3. Se abrirá un panel lateral con la hora actual
4. **Selecciona una celda** en la hoja
5. **Elige el formato** de hora (opcional)
6. **Haz clic** en "Insertar Hora" o "Insertar Fecha + Hora"
7. ¡La hora se insertará en la celda! ✅

---

## 🔧 Personalización

### Cambiar el ícono del complemento:

En `manifest.xml`, agregar:
```xml
<Hosts>
    <Host Name="Workbook">
        <DesktopFormFactor>
            <GetStarted>
                <Title resid="GetStartedTitle"/>
                <Description resid="GetStartedDescription"/>
                <LearnMoreUrl DefaultValue="https://aka.ms/officedeveloper"/>
            </GetStarted>
            <FunctionFile resid="Commands.Function.Url"/>
            <ExtensionPoint xsi:type="PrimaryCommandSurface">
                <OfficeMenu id="ContextMenuCell">
                    <Control id="InsertarHora" xsi:type="Button">
                        <Label resid="ButtonLabel"/>
                        <Subroutine FunctionName="insertarHora"/>
                    </Control>
                </OfficeMenu>
            </ExtensionPoint>
        </DesktopFormFactor>
    </Host>
</Hosts>
```

### Cambiar los colores:

En `taskpane.css`, busca y modifica:
```css
.btn-primary {
    background-color: #0078d4;  /* Azul principal - cambiar aquí */
}

.info-section {
    background: linear-gradient(135deg, #0078d4 0%, #107c10 100%);  /* Gradiente */
}
```

---

## 🐛 Solución de Problemas

### "El complemento no carga"
- ✅ Verifica que la URL en manifest.xml sea correcta
- ✅ Asegúrate de usar HTTPS
- ✅ Comprueba que los archivos sean accesibles
- ✅ Usa las herramientas de desarrollador (F12) para ver errores

### "Error: No se puede ejecutar el macro"
- ✅ Esto es normal en Excel Web (no hay VBA)
- ✅ Este complemento usa Office JS API (más moderno)

### "La celda no se actualiza"
- ✅ Recarga la página (Ctrl + F5)
- ✅ Verifica permisos del archivo

### "No veo el complemento después de instalarlo"
- ✅ Ve a **Insertar** > **Complementos** > **Mis complementos cargados**
- ✅ Busca el complemento en la lista
- ✅ Reinicia Excel Online

---

## 📐 Estructura técnica

```
Office.onReady()
    ↓
Inicializar complemento
    ↓
Excel.run(async (context) => {
    activeCell = context.application.activeCell
    cell.values = [[ hora ]]
    context.sync()
})
```

---

## 🔐 Seguridad

- El complemento solo tiene permisos de **lectura/escritura** en el documento
- No accede a datos personales del usuario
- No se conecta a servidores externos
- Cumple con estándares de Office 365

---

## 📚 Recursos Adicionales

### Documentación oficial:
- [Office Add-ins Documentation](https://docs.microsoft.com/es-es/office/dev/add-ins/)
- [Excel JavaScript API](https://docs.microsoft.com/es-es/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
- [Office Add-ins Best Practices](https://docs.microsoft.com/es-es/office/dev/add-ins/overview/best-practices-for-developing-office-add-ins)

### Herramientas útiles:
- [Visual Studio Code](https://code.visualstudio.com/)
- [Office Add-ins Validator](https://validator.oapps.office.net/)
- [Yeoman Generator for Office Add-ins](https://github.com/OfficeDev/generator-office)

---

## 📝 Ejemplo de Implementación con GitHub Pages

1. **Crea un repositorio** en GitHub
2. **Copia los archivos** a la rama `main`
3. **Habilita GitHub Pages** (Settings > Pages > Deploy from branch > main)
4. **Espera a que se publique** (unos segundos)
5. Tu URL será: `https://tunombre.github.io/nombre-repo/`
6. En manifest.xml:
   ```xml
   <SourceLocation DefaultValue="https://tunombre.github.io/nombre-repo/taskpane.html"/>
   ```

---

## ✅ Checklist Final

- [ ] Archivos descargados
- [ ] Servidor web configurado (o GitHub Pages)
- [ ] manifest.xml actualizado con URL correcta
- [ ] Archivos HTML, CSS, JS en el servidor
- [ ] URL es HTTPS
- [ ] Probado en Excel Online
- [ ] Complemento cargado exitosamente

---

## 🎉 ¡Listo!

Tu complemento de Excel Web está funcionando. Ahora puedes:
- Usarlo en Excel Online
- Compartirlo con tu equipo
- Extenderlo con más funcionalidades
- Publicarlo en la Office Store (opcional)

¿Preguntas? Consulta la documentación oficial de Microsoft.
