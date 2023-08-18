# Aplicación de Consola de Windows - Respaldos(Transmisión de datos) mediante FTP

**Autor:**

- Jorge Daniel Velasco García @J0rgeV

Esta aplicación es capaz de realizar envío de información a través del FTP a un servidor. La misma puede realizar distintos envíos a diferentes locaciones o servidores con distintos parámetros que igualmente se define con anterioridad.

Por ejemplo, si queremos enviar información al Servidor01 pasándole como parámetro ```EnviaServ01```, será pertinente modificar el contenido donde se obtiene parámetros y así determinar el nombre del archivo/carpeta a subir, el destino, con qué credenciales se realizará la tranferencia, etc.
Ya una vez realizado esta acción solamente será ejecutar la tarea con el siguiente comando en terminal:

```
RespaldoFTP.exe EnviaServ01
```

**Compilación y ejecución:**

Se requiere del software Visual Studio con el objetivo de modificar el código de manera más sencilla y lograr que funcione. Para ello está el [siguiente enlace](https://learn.microsoft.com/es-es/visualstudio/install/install-visual-studio?view=vs-2022) como guía.

Una vez modificado las líneas correspondientes, es decir, los parámetros y locaciones del origen y destino de la transmisión, en el mismo software realizamos el proceso de compilación. 

Así, para la ejecución se requiere de abrir una terminal en la carpeta donde se encuentra el ejecutable de la aplicación y colocar lo siguiente:

```
RespaldoFTP.exe Parámetro01
```