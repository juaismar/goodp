# GOODP - Generador de OpenDocument Presentation en Go

GOODP es una biblioteca en Go para crear presentaciones en formato OpenDocument (.odp) de manera programática. Permite crear diapositivas, añadir texto, imágenes y establecer fondos personalizados.

## Instalación

```bash
go get github.com/juaismar/goodp
```

## Características

- Creación de presentaciones en formato ODP
- Soporte para diferentes relaciones de aspecto (16:9, 4:3)
- Añadir texto con estilos personalizados
- Insertar imágenes
- Establecer fondos (imágenes o colores sólidos)
- Personalización de tamaños de diapositiva

## Uso Básico

```go
package main

import (
    "log"
    "os"
    "goodp"
)

func main() {
    // Crear una nueva presentación (por defecto 16:9)
    presentacion := goodp.New()

    // Añadir una diapositiva con título
    presentacion.AddSlide("Mi Primera Presentación", "Contenido de ejemplo")

    // Guardar la presentación
    if err := presentacion.Save("mi_presentacion.odp"); err != nil {
        log.Fatal(err)
    }
}
```

## Ejemplos de Uso

### Configurar el Tamaño de la Presentación

```go
// Usar relación de aspecto 16:9 (por defecto)
presentacion.SetSlideSize(goodp.AspectRatio169)

// Usar relación de aspecto 4:3
presentacion.SetSlideSize(goodp.AspectRatio43)

// O establecer un tamaño personalizado (en centímetros)
presentacion.SetCustomSlideSize(25.4, 19.05)
```

### Añadir Texto con Estilo

```go
// Obtener referencia a la diapositiva
slide := presentacion.AddBlankSlide()

// Establecer estilo de texto (tamaño, fuente, color, negrita, cursiva)
presentacion.SetTextStyle(slide, 24, "Arial", "#FF0000", true, false)

// Añadir cuadro de texto (slide, contenido, x, y, ancho, alto en cm)
presentacion.AddTextBox(slide, "Texto con estilo", 2, 2, 10, 2)
```

### Insertar Imágenes

```go
// Leer imagen
imageData, err := os.ReadFile("imagen.png")
if err != nil {
    log.Fatal(err)
}

// Obtener referencia a la diapositiva
slide := presentacion.AddBlankSlide()

// Añadir imagen (slide, datos, extensión, x, y, ancho, alto en cm)
if err := presentacion.AddImage(slide, imageData, ".png", 15, 5, 10, 8); err != nil {
    log.Fatal(err)
}
```

### Establecer Fondos

```go
// Establecer fondo de color para toda la presentación
presentacion.SetBackgroundColor("#FF0000")

// O establecer una imagen de fondo global
imageData, err := os.ReadFile("fondo.jpg")
if err != nil {
    log.Fatal(err)
}
if err := presentacion.SetBackgroundImage(imageData, ".jpg"); err != nil {
    log.Fatal(err)
}

// Establecer fondo solo para una diapositiva específica
slide := presentacion.AddBlankSlide()
if err := presentacion.SetSlideBackground(slide, imageData, ".jpg"); err != nil {
    log.Fatal(err)
}

// O un color de fondo específico para una diapositiva
if err := presentacion.SetSlideBackgroundColor(slide, "#0000FF"); err != nil {
    log.Fatal(err)
}
```

## Ejemplo Completo

Puedes encontrar un ejemplo completo en el archivo [ejemplo_uso.go](example/ejemplo_uso.go).

## Limitaciones

- Solo soporta formatos de imagen comunes (PNG, JPEG, etc.)
- No soporta animaciones ni transiciones
- No soporta edición de presentaciones existentes

## Contribuir

Las contribuciones son bienvenidas. Por favor, abre un issue para discutir los cambios que te gustaría hacer.

## Licencia

[MIT](LICENSE)