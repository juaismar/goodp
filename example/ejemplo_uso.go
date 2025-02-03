package main

import (
	"log"
	"os"

	"goodp"
)

func main() {
	// Crear una nueva presentación (por defecto será 16:9)
	presentacion := goodp.New()

	// O puedes establecer explícitamente el tamaño
	// presentacion.SetSlideSize(goodp.AspectRatio169) // 16:9
	// presentacion.SetSlideSize(goodp.AspectRatio43)  // 4:3
	// O usar un tamaño personalizado (en centímetros)
	// presentacion.SetCustomSlideSize(25.4, 14.29)

	// Añadir una diapositiva con título
	presentacion.AddSlide("Primera página", "")

	// Leer la imagen
	imageData, err := os.ReadFile("./img/grafica1.png")
	if err != nil {
		log.Fatal(err)
	}

	// Añadir la imagen (usando la extensión del archivo)
	err = presentacion.AddImage(imageData, ".png", 15, 5, 10, 8)
	if err != nil {
		log.Fatal(err)
	}

	// Establecer fondo global
	imageData, err = os.ReadFile("./img/gradient-background.jpg")
	if err != nil {
		log.Fatal(err)
	}
	err = presentacion.SetBackgroundImage(imageData, ".jpg")
	if err != nil {
		log.Fatal(err)
	}

	// Añadir diapositiva y establecer su fondo específico
	presentacion.AddBlankSlide()
	imageData, err = os.ReadFile("./img/light-background.jpg")
	if err != nil {
		log.Fatal(err)
	}
	err = presentacion.SetCurrentSlideBackground(imageData, ".jpg")
	if err != nil {
		log.Fatal(err)
	}

	// Añadir cuadros de texto en diferentes posiciones
	// Los parámetros son: contenido, x, y, ancho, alto (en centímetros)
	presentacion.AddTextBox("Texto en la esquina superior", 2, 2, 10, 2)

	presentacion.SetTextStyle(24, "Arial", "#FF0000", true, false) // Rojo, negrita
	presentacion.AddTextBox("Texto en el centro", 10, 8, 8, 4)

	presentacion.SetTextStyle(18, "Times New Roman", "#0000FF", false, true) // Azul, cursiva
	presentacion.AddTextBox("Texto en la parte inferior", 15, 15, 10, 3)

	presentacion.AddSlide("Tercera página", "Your name is 宮水 三葉")

	// Guardar la presentación
	err = presentacion.Save("Presentacion.odp")
	if err != nil {
		log.Fatal(err)
	}
}
