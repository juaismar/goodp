# GOODP - Generador de OpenDocument Presentation en Go

GOODP es una biblioteca en Go para crear presentaciones en formato OpenDocument (.odp) de manera programática. Permite crear diapositivas, añadir texto, imágenes y establecer fondos personalizados.

## Instalación

```bash
go get github.com/tu-usuario/goodp
```

## Características

- Creación de presentaciones en formato ODP
- Soporte para diferentes relaciones de aspecto (16:9, 4:3)
- Añadir texto con estilos personalizados
- Insertar imágenes
- Establecer fondos (imágenes o colores sólidos)
- Personalización de tamaños de diapositiva

## Uso Básico

```
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
err := presentacion.Save("mi_presentacion.odp")
if err != nil {
log.Fatal(err)
}
}
```

## Ejemplos de Uso

### Configurar el Tamaño de la Presentación

```
// Usar relación de aspecto 16:9 (por defecto)
presentacion.SetSlideSize(goodp.AspectRatio169)
// Usar relación de aspecto 4:3
presentacion.SetSlideSize(goodp.AspectRatio43)
// O establecer un tamaño personalizado (en centímetros)
presentacion.SetCustomSlideSize(25.4, 19.05)
```

### Añadir Texto con Estilo
```
// Establecer estilo de texto (tamaño, fuente, color, negrita, cursiva)
presentacion.SetTextStyle(24, "Arial", "#FF0000", true, false)
// Añadir cuadro de texto (contenido, x, y, ancho, alto en cm)
presentacion.AddTextBox("Texto con estilo", 2, 2, 10, 2)
```
### Insertar Imágenes
```
// Leer imagen
imageData, err := os.ReadFile("imagen.png")
if err != nil {
log.Fatal(err)
}
// Añadir imagen (datos, extensión, x, y, ancho, alto en cm)
err = presentacion.AddImage(imageData, ".png", 15, 5, 10, 8)
if err != nil {
log.Fatal(err)
}
```
### Establecer Fondos
```
// Establecer fondo de color para toda la presentación
presentacion.SetBackgroundColor("#FF0000")
// O establecer una imagen de fondo
imageData, err := os.ReadFile("fondo.jpg")
if err != nil {
log.Fatal(err)
}
err = presentacion.SetBackgroundImage(imageData, ".jpg")
if err != nil {
log.Fatal(err)
}
// Establecer fondo solo para la diapositiva actual
err = presentacion.SetCurrentSlideBackground(imageData, ".jpg")
// O un color de fondo
err = presentacion.SetCurrentSlideBackgroundColor("#0000FF")
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