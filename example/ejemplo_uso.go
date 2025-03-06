package main

import (
	"log"
	"os"

	"github.com/juaismar/goodp"
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
	silde1 := presentacion.AddSlide("Primera página", "")

	// Leer la imagen
	imageData, err := os.ReadFile("./img/grafica1.png")
	if err != nil {
		log.Fatal(err)
	}

	// Añadir la imagen (usando la extensión del archivo)
	err = presentacion.AddImage(silde1, imageData, ".png", 15, 5, 10, 8)
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

	// Establecer un fondo de color global
	//presentacion.SetBackgroundColor("#FF0000") // Fondo rojo para toda la presentación

	// Añadir diapositiva y establecer su fondo específico
	silde2 := presentacion.AddBlankSlide()
	imageData, err = os.ReadFile("./img/light-background.jpg")
	if err != nil {
		log.Fatal(err)
	}
	err = presentacion.SetSlideBackground(silde2, imageData, ".jpg")
	if err != nil {
		log.Fatal(err)
	}

	// Añadir cuadros de texto en diferentes posiciones
	// Los parámetros son: contenido, x, y, ancho, alto (en centímetros)
	presentacion.SetTextStyle(silde2, 20, "Arial", "#000000", false, false)
	presentacion.AddTextBox(silde2, "Texto en la esquina superior", 2, 2, 10, 2,
		&goodp.TextProperties{
			HorizontalAlign: "left",
			VerticalAlign:   "top",
		})

	presentacion.SetTextStyle(silde2, 24, "Arial", "#FF0000", true, false) // Rojo, negrita
	presentacion.AddTextBox(silde2, "Texto en el centro", 10, 8, 8, 4,
		&goodp.TextProperties{
			HorizontalAlign: "center",
			VerticalAlign:   "middle",
		})

	presentacion.SetTextStyle(silde2, 18, "Times New Roman", "#0000FF", false, true) // Azul, cursiva
	presentacion.AddTextBox(silde2, "Texto en la parte inferior", 15, 15, 10, 3,
		&goodp.TextProperties{
			HorizontalAlign: "right",
			VerticalAlign:   "bottom",
		})
	presentacion.AddTextBox(silde2, "Texto justificado", 2, 15, 10, 3,
		&goodp.TextProperties{
			HorizontalAlign: "justify",
			VerticalAlign:   "middle",
		})

	presentacion.AddSlide("Tercera página", "Your name is 宮水 三葉")

	silde4 := presentacion.AddSlide("Cuarta página", "")
	err = presentacion.SetSlideBackgroundColor(silde4, "#0000FF") // Fondo azul para esta diapositiva
	if err != nil {
		log.Fatal(err)
	}
	presentacion.AddTextBox(silde4, "Texto sin props", 2, 15, 10, 3, nil)

	// Guardar la presentación
	err = presentacion.Save("Presentacion.odp")
	if err != nil {
		log.Fatal(err)
	}
}
