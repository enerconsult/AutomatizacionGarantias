# Guía de Automatización con Programador de Tareas

Para ejecutar el script `run_daily.py` todos los días automáticamente, siga estos pasos en Windows:

1.  Abra el **Programador de Tareas** (Task Scheduler). Puede buscarlo en el menú Inicio.
2.  Haga clic en **"Crear tarea básica..."** en el panel derecho.
3.  **Nombre**: Coloque un nombre, ej: `Descarga Automática XM`.
4.  **Desencadenador**: Seleccione **"Diariamente"**.
5.  **Día/Hora**: Configure la hora deseada (ej: 08:00 AM) y "Repetir cada: 1 día".
6.  **Acción**: Seleccione **"Iniciar un programa"**.
7.  **Programa o script**:
    *   Examine y seleccione su ejecutable de Python.
    *   Generalmente está en: `C:\Users\jqele\anaconda3\python.exe` (según su configuración previa).
    *   Si no sabe dónde está, ejecute `where python` en una terminal.
8.  **Argumentos (Opcional)**: Escriba el nombre del script entre comillas:
    *   `"g:\Mi unidad\Garantías\Automatización\run_daily.py"`
9.  **Iniciar en (Opcional)**: ¡IMPORTANTE! Coloque la ruta de la carpeta donde está el script:
    *   `g:\Mi unidad\Garantías\Automatización\`
    *   *Nota: Esto asegura que las descargas se guarden en la carpeta correcta.*
10. Finalice el asistente.

## Prueba
Haga clic derecho sobre la tarea creada y seleccione **"Ejecutar"** para verificar que funciona correctamente. Debería aparecer momentáneamente una ventana negra o ejecutarse en segundo plano (según configuración de usuario).
