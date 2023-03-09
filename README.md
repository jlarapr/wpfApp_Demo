# wpfApp_Demo
Demo Wpf

## By Jose Lara
2023-02-05

### Objetivo

# Crear un archivo de Excel, que contenga los totales de Imágenes  y los totales de reclamaciones
    - Imágenes es igual a todos los record de la Base de Datos 
    - Reclamaciones es igual a todos los record que en la columna V1PAGE sea diferente 99

# Ejemplo del Out Put Final (Totales.xlsx)

|  Date    | Total Imágenes | Total Reclamaciones |
|:--------:|:--------------:|:-------------------:|
| 3/9/2023 |1,000           |500                  |

# Note:
    - La base de datos se encuentra en el directorio assets/vdeFiles
    - Nombre de la base de datos (SSS1503.DBF)
    - Package EPPlus para poder crear un archivo de excel (https://www.nuget.org/packages/EPPlus/)
    - .Net 6
    - WPF
    - MVVM

  

# Estarían programado en el Botón Save
     - Código del Botón se encuentra en (main_vm.cs)
# UI (User interface)

![alt text](https://github.com/jlarapr/wpfApp_Demo/blob/dev/assets/main.png?raw=true)


# Estructura de directorios
1. assets
    - vdeFiles
    - app.ico
2. src
    - core
        - data
            - configuration
            - repository
        - interfaces
            - repository
            - service
        - models
        - services
        - viewModels
    - ui
        - views

# assets
    vdeFiles
        base de dato de ejemplo para, asi poder hacer el demo. 

# src
    contiene el código de la aplicaci�n 

# core
    contiene todas la programación de la aplicaci�n
    data, interface, models, services y viewModels

# ui (user interface)
    contiene todos los archivos de las interfase del usuario (*.xaml)   

