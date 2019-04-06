#Utilidades PowerShell

Utilidades desarrolladas en PowerShell.

# Instalación
Para poder ejecutar los PowerShell es preciso tener instalado NuGet y los modulos necesarios para su ejecución. Si es la primera vez que ejecuta el programa, le solicitará permiso para instalar NuGet y los módulos que precise. Acepte para poder ejecutarlos.

```
Se necesita el proveedor de NuGet para continuar
PowerShellGet necesita la versión del proveedor de NuGet '2.8.5.201' o posterior para interactuar con repositorios basados en NuGet. El proveedor de NuGet debe estar disponible en
'C:\Program Files\PackageManagement\ProviderAssemblies' o 'C:\Users\Oscar\AppData\Local\PackageManagement\ProviderAssemblies'. También puedes instalar el proveedor de NuGet ejecutando 'Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force'. ¿Quieres que PowerShellGet se instale e importe el proveedor de NuGet ahora?
[S] Sí [N] No [U] Suspender [?] Help (default is "Sí"):
```
Pulsar [Enter]
S
```
Estás instalando los módulos desde un repositorio que no es de confianza. Si confías en este repositorio, cambia su valor InstallationPolicy ejecutando el cmdlet Set-PSRepository.
¿Estás seguro de que quieres instalar los módulos de 'PSGallery'?
[S] Sí [O] Sí a todo [N] No [T] No a todo [U] Suspender [?] Help (default is "No"): 

Pulsar [Enter (Sí)]

# Utilidades

## excel-mail-remover
A partir de una excel con una columna de emails y otra excel con los emails a eliminar, busca los emails y los elimina (ignorando mayúsculas/minusculas).
