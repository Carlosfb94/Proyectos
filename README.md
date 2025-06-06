# Proyectos
Proyectos de prueba para desarrollar cosas.

## Shipping Tracker

El repositorio incluye `shipping_tracker.py`, un script para leer
manifiestos de envío en PDF o Excel y actualizar un archivo Excel con
el estado de cada orden.  La lógica proviene de un script más grande y
se ha refactorizado para ser más modular y fácil de mantener.

Ejemplo de uso:

```bash
python shipping_tracker.py carpeta_con_archivos --excel resultados.xlsx
```

La clave de la API de 2Captcha se toma del entorno `API_KEY_2CAPTCHA` si
se requieren consultas a Cruz del Sur.
