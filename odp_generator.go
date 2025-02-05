package goodp

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
	"text/template"
)

// TODO Refactor this to use a list of sizes and override with a custom size
// Constantes para los tipos de diapositivas comunes
const (
	AspectRatio169 = "16:9"
	AspectRatio43  = "4:3"
)

// SlideSize representa las dimensiones de la diapositiva
type SlideSize struct {
	Width  float64
	Height float64
}

// Tamaños predefinidos (en centímetros)
var (
	defaultSize169 = SlideSize{Width: 33.867, Height: 19.05} // 16:9 (equivalente a 1920x1080 en cm)
	defaultSize43  = SlideSize{Width: 25.4, Height: 19.05}   // 4:3 (equivalente a 1024x768 en cm)
)

// Modificar la estructura BackgroundImage para soportar diferentes tipos de fondo
type BackgroundType int

const (
	BackgroundImage BackgroundType = iota
	BackgroundColor
)

type Background struct {
	Type  BackgroundType
	Data  []byte // Para imágenes
	Name  string // Para imágenes
	Color string // Para colores sólidos
}

type ODPGenerator struct {
	Slides     []Slide
	SlideSize  SlideSize
	Background *Background
}

type Slide struct {
	Title        string
	Content      string
	TextBoxes    []TextBox
	Images       []Image
	currentStyle TextStyle
	Background   *Background
}

type TextBox struct {
	Content string
	X       string // Posición X en cm
	Y       string // Posición Y en cm
	Width   string // Ancho en cm
	Height  string // Alto en cm
	Style   TextStyle
}

type Image struct {
	Data   []byte // Cambiamos Path por Data para almacenar los bytes de la imagen
	X      string
	Y      string
	Width  string
	Height string
	Name   string
}

type TextStyle struct {
	FontSize   string
	FontFamily string
	Color      string
	Bold       bool
	Italic     bool
}

// New crea una nueva instancia de ODPGenerator con tamaño 16:9 por defecto
func New() *ODPGenerator {
	return &ODPGenerator{
		Slides:    make([]Slide, 0),
		SlideSize: defaultSize169,
	}
}

// SetSlideSize establece el tamaño de las diapositivas
func (g *ODPGenerator) SetSlideSize(aspectRatio string) {
	switch aspectRatio {
	case AspectRatio169:
		g.SlideSize = defaultSize169
	case AspectRatio43:
		g.SlideSize = defaultSize43
	default:
		// Si no se reconoce el aspect ratio, usar 16:9 por defecto
		g.SlideSize = defaultSize169
	}
}

// SetCustomSlideSize establece un tamaño personalizado para las diapositivas
func (g *ODPGenerator) SetCustomSlideSize(width, height float64) {
	g.SlideSize = SlideSize{
		Width:  width,
		Height: height,
	}
}

// AddSlide añade una nueva diapositiva a la presentación y devuelve un puntero a ella
func (g *ODPGenerator) AddSlide(title string, content string) *Slide {
	g.Slides = append(g.Slides, Slide{
		Title:   title,
		Content: content,
	})
	return &g.Slides[len(g.Slides)-1]
}

// AddBlankSlide añade una diapositiva en blanco a la presentación y devuelve un puntero a ella
func (g *ODPGenerator) AddBlankSlide() *Slide {
	g.Slides = append(g.Slides, Slide{
		Title:   "",
		Content: "",
	})
	return &g.Slides[len(g.Slides)-1]
}

// SetTextStyle establece el estilo para el próximo texto que se añada
func (g *ODPGenerator) SetTextStyle(slide *Slide, fontSize float64, fontFamily, color string, bold, italic bool) {
	if len(g.Slides) == 0 {
		g.AddBlankSlide()
	}

	slide.currentStyle = TextStyle{
		FontSize:   fmt.Sprintf("%.2fpt", fontSize),
		FontFamily: fontFamily,
		Color:      color,
		Bold:       bold,
		Italic:     italic,
	}
}

// AddTextBox añade un cuadro de texto a la última diapositiva
func (g *ODPGenerator) AddTextBox(slide *Slide, content string, x, y, width, height float64) {
	if len(g.Slides) == 0 {
		g.AddBlankSlide()
	}

	slide.TextBoxes = append(slide.TextBoxes, TextBox{
		Content: content,
		X:       fmt.Sprintf("%.2fcm", x),
		Y:       fmt.Sprintf("%.2fcm", y),
		Width:   fmt.Sprintf("%.2fcm", width),
		Height:  fmt.Sprintf("%.2fcm", height),
		Style:   slide.currentStyle, // Usar el estilo actual
	})
}

// AddImage añade una imagen a la diapositiva especificada.
// El parámetro extension debe incluir el punto (por ejemplo: ".jpg", ".png")
func (g *ODPGenerator) AddImage(slide *Slide, imageData []byte, extension string, x, y, width, height float64) error {
	// Validar que el slide pertenece a esta presentación
	slideIndex := -1
	for i := range g.Slides {
		if &g.Slides[i] == slide {
			slideIndex = i
			break
		}
	}
	if slideIndex == -1 {
		return fmt.Errorf("la diapositiva especificada no pertenece a esta presentación")
	}

	// Validar la extensión
	extension = strings.ToLower(strings.TrimSpace(extension))
	if !strings.HasPrefix(extension, ".") {
		extension = "." + extension
	}

	// Validar que sea una extensión de imagen soportada
	validExtensions := map[string]bool{
		".jpg": true, ".jpeg": true,
		".png": true,
		".gif": true,
		".bmp": true,
		".svg": true,
	}
	if !validExtensions[extension] {
		return fmt.Errorf("formato de imagen no soportado: %s", extension)
	}

	// Validar que imageData no esté vacío
	if len(imageData) == 0 {
		return fmt.Errorf("los datos de la imagen están vacíos")
	}

	// Validar dimensiones
	if width <= 0 || height <= 0 {
		return fmt.Errorf("las dimensiones de la imagen deben ser positivas")
	}

	// Validar que la imagen cabe en la diapositiva
	if x < 0 || y < 0 ||
		x+width > g.SlideSize.Width ||
		y+height > g.SlideSize.Height {
		return fmt.Errorf("la imagen se sale de los límites de la diapositiva (%.2f x %.2f)",
			g.SlideSize.Width, g.SlideSize.Height)
	}

	// Generar un nombre único para la imagen
	imageName := fmt.Sprintf("Pictures/slide%d_image%d%s",
		slideIndex,
		len(slide.Images),
		extension)

	slide.Images = append(slide.Images, Image{
		Data:   imageData,
		X:      fmt.Sprintf("%.2fcm", x),
		Y:      fmt.Sprintf("%.2fcm", y),
		Width:  fmt.Sprintf("%.2fcm", width),
		Height: fmt.Sprintf("%.2fcm", height),
		Name:   imageName,
	})

	return nil
}

// SetBackgroundImage establece una imagen de fondo para todas las diapositivas
func (g *ODPGenerator) SetBackgroundImage(imageData []byte, extension string) error {
	imageName := fmt.Sprintf("media/background.%s", strings.ToLower(strings.TrimPrefix(extension, ".")))

	g.Background = &Background{
		Type: BackgroundImage,
		Data: imageData,
		Name: imageName,
	}

	return nil
}

// SetSlideBackground establece una imagen de fondo para la diapositiva especificada.
// El parámetro extension debe incluir el punto (por ejemplo: ".jpg", ".png")
func (g *ODPGenerator) SetSlideBackground(slide *Slide, imageData []byte, extension string) error {
	// Validar que el slide pertenece a esta presentación
	slideIndex := -1
	for i := range g.Slides {
		if &g.Slides[i] == slide {
			slideIndex = i
			break
		}
	}
	if slideIndex == -1 {
		return fmt.Errorf("la diapositiva especificada no pertenece a esta presentación")
	}

	// Validar la extensión
	extension = strings.ToLower(strings.TrimSpace(extension))
	if !strings.HasPrefix(extension, ".") {
		extension = "." + extension
	}

	// Validar que sea una extensión de imagen soportada
	validExtensions := map[string]bool{
		".jpg": true, ".jpeg": true,
		".png": true,
		".gif": true,
		".bmp": true,
		".svg": true,
	}
	if !validExtensions[extension] {
		return fmt.Errorf("formato de imagen no soportado: %s", extension)
	}

	// Validar que imageData no esté vacío
	if len(imageData) == 0 {
		return fmt.Errorf("los datos de la imagen están vacíos")
	}

	imageName := fmt.Sprintf("media/slide%d_background%s", slideIndex, extension)

	slide.Background = &Background{
		Type: BackgroundImage,
		Data: imageData,
		Name: imageName,
	}

	return nil
}

// SetBackgroundColor establece un fondo de color para la presentación
func (g *ODPGenerator) SetBackgroundColor(color string) {
	g.Background = &Background{
		Type:  BackgroundColor,
		Color: color,
	}
}

// SetSlideBackgroundColor establece un color de fondo para la diapositiva especificada.
// El color debe estar en formato hexadecimal (#RRGGBB) o ser un nombre de color válido.
func (g *ODPGenerator) SetSlideBackgroundColor(slide *Slide, color string) error {
	// Validar que el slide pertenece a esta presentación
	slideFound := false
	for i := range g.Slides {
		if &g.Slides[i] == slide {
			slideFound = true
			break
		}
	}
	if !slideFound {
		return fmt.Errorf("la diapositiva especificada no pertenece a esta presentación")
	}

	// Validar el formato del color
	color = strings.TrimSpace(color)
	if !strings.HasPrefix(color, "#") {
		color = "#" + color
	}
	if len(color) != 7 {
		return fmt.Errorf("formato de color inválido: debe ser #RRGGBB")
	}

	// Validar que los caracteres sean hexadecimales válidos
	for _, c := range color[1:] {
		if !((c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F')) {
			return fmt.Errorf("formato de color inválido: caracteres no hexadecimales")
		}
	}

	slide.Background = &Background{
		Type:  BackgroundColor,
		Color: color,
	}

	return nil
}

// SaveStream genera y devuelve los bytes del archivo ODP
func (g *ODPGenerator) SaveStream() ([]byte, error) {
	// Crear el archivo ZIP (ODP es un archivo ZIP)
	buf := new(bytes.Buffer)
	zipWriter := zip.NewWriter(buf)

	// Añadir mimetype
	mimetypeWriter, err := zipWriter.Create("mimetype")
	if err != nil {
		return nil, err
	}
	_, err = mimetypeWriter.Write([]byte("application/vnd.oasis.opendocument.presentation"))
	if err != nil {
		return nil, err
	}

	// Añadir content.xml
	contentWriter, err := zipWriter.Create("content.xml")
	if err != nil {
		return nil, err
	}
	err = g.writeContent(contentWriter)
	if err != nil {
		return nil, err
	}

	// Añadir styles.xml
	stylesWriter, err := zipWriter.Create("styles.xml")
	if err != nil {
		return nil, err
	}
	err = g.writeStyles(stylesWriter)
	if err != nil {
		return nil, err
	}

	// Añadir settings.xml
	settingsWriter, err := zipWriter.Create("settings.xml")
	if err != nil {
		return nil, err
	}
	err = g.writeSettings(settingsWriter)
	if err != nil {
		return nil, err
	}

	// Añadir configurations2/accelerator/current.xml
	configWriter, err := zipWriter.Create("configurations2/accelerator/current.xml")
	if err != nil {
		return nil, err
	}
	err = g.writeConfigurations(configWriter)
	if err != nil {
		return nil, err
	}

	// Añadir manifest
	manifestWriter, err := zipWriter.Create("META-INF/manifest.xml")
	if err != nil {
		return nil, err
	}
	err = g.writeManifest(manifestWriter)
	if err != nil {
		return nil, err
	}

	// Añadir la imagen de fondo global si existe y es una imagen
	if g.Background != nil && g.Background.Type == BackgroundImage {
		imageWriter, err := zipWriter.Create(g.Background.Name)
		if err != nil {
			return nil, err
		}
		_, err = imageWriter.Write(g.Background.Data)
		if err != nil {
			return nil, err
		}
	}

	// Añadir las imágenes de fondo por diapositiva
	for _, slide := range g.Slides {
		if slide.Background != nil && slide.Background.Type == BackgroundImage {
			imageWriter, err := zipWriter.Create(slide.Background.Name)
			if err != nil {
				return nil, err
			}
			_, err = imageWriter.Write(slide.Background.Data)
			if err != nil {
				return nil, err
			}
		}
	}

	// Añadir las imágenes al archivo ZIP
	for _, slide := range g.Slides {
		for _, img := range slide.Images {
			imageWriter, err := zipWriter.Create(img.Name)
			if err != nil {
				return nil, err
			}

			_, err = imageWriter.Write(img.Data)
			if err != nil {
				return nil, err
			}
		}
	}

	// Cerrar el ZIP
	err = zipWriter.Close()
	if err != nil {
		return nil, err
	}

	return buf.Bytes(), nil
}

// Modificar Save para usar SaveStream
func (g *ODPGenerator) Save(filename string) error {
	if !strings.HasSuffix(filename, ".odp") {
		filename += ".odp"
	}

	// Obtener los bytes usando SaveStream
	data, err := g.SaveStream()
	if err != nil {
		return err
	}

	// Escribir el archivo
	return os.WriteFile(filename, data, 0644)
}

func (g *ODPGenerator) writeContent(writer io.Writer) error {
	const contentTemplate = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content 
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
    xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
    xmlns:xlink="http://www.w3.org/1999/xlink"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
    xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
    xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"
    xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
    xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"
    xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0"
    xmlns:math="http://www.w3.org/1998/Math/MathML"
    xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0"
    xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"
    xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"
    xmlns:ooo="http://openoffice.org/2004/office"
    xmlns:ooow="http://openoffice.org/2004/writer"
    xmlns:oooc="http://openoffice.org/2004/calc"
    xmlns:dom="http://www.w3.org/2001/xml-events"
    xmlns:xforms="http://www.w3.org/2002/xforms"
    xmlns:xsd="http://www.w3.org/2001/XMLSchema"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    office:version="1.2">
    <office:scripts/>
    <office:font-face-decls/>
    <office:automatic-styles>
        {{if .Background}}
        <style:style style:family="drawing-page" style:name="backgroundStyle">
            <style:drawing-page-properties 
                {{if eq .Background.Type 0}}
                draw:fill="bitmap" 
                draw:fill-image-name="backgroundImage" 
                style:repeat="stretch"
                draw:background-size="border" 
                {{else}}
                draw:fill="solid"
                draw:fill-color="{{.Background.Color}}"
                {{end}}
                presentation:background-objects-visible="true" 
                presentation:background-visible="false"
                presentation:display-header="false" 
                presentation:display-footer="false" 
                presentation:display-page-number="false" 
                presentation:display-date-time="false"/>
        </style:style>
        {{end}}
        {{range $index, $slide := .Slides}}
            {{if $slide.Background}}
            <style:style style:family="drawing-page" style:name="slideBackground{{$index}}">
                <style:drawing-page-properties 
                    {{if eq $slide.Background.Type 0}}
                    draw:fill="bitmap" 
                    draw:fill-image-name="slideBackground{{$index}}" 
                    style:repeat="stretch"
                    draw:background-size="border" 
                    {{else}}
                    draw:fill="solid"
                    draw:fill-color="{{$slide.Background.Color}}"
                    {{end}}
                    presentation:background-objects-visible="true" 
                    presentation:background-visible="false"
                    presentation:display-header="false" 
                    presentation:display-footer="false" 
                    presentation:display-page-number="false" 
                    presentation:display-date-time="false"/>
            </style:style>
            {{end}}
        {{end}}
        <style:style style:name="dp1" style:family="drawing-page">
            <style:drawing-page-properties presentation:background-visible="true"
                                         presentation:background-objects-visible="true"
                                         presentation:display-footer="true"
                                         presentation:display-page-number="false"
                                         presentation:display-date-time="true"/>
        </style:style>
        <style:style style:name="gr1" style:family="graphic">
            <style:graphic-properties draw:stroke="none" draw:fill="none" draw:textarea-horizontal-align="center"/>
        </style:style>
        <style:style style:name="gr2" style:family="graphic">
            <style:graphic-properties draw:stroke="none" draw:fill="none"/>
        </style:style>
        <style:style style:name="P1" style:family="paragraph">
            <style:paragraph-properties fo:text-align="center"/>
        </style:style>
        <style:style style:name="P2" style:family="paragraph">
            <style:paragraph-properties fo:text-align="left"/>
        </style:style>
    </office:automatic-styles>
    <office:body>
        <office:presentation>
            {{range $index, $slide := .Slides}}
            <draw:page draw:name="page{{$index}}" 
                      {{if $slide.Background}}
                      draw:style-name="slideBackground{{$index}}"
                      {{else if $.Background}}
                      draw:style-name="backgroundStyle"
                      {{else}}
                      draw:style-name="dp1"
                      {{end}}
                      draw:master-page-name="Default">
                {{if .Title}}
                <draw:frame draw:style-name="gr1" draw:text-style-name="P1" draw:layer="layout" 
                           svg:width="{{printf "%.2fcm" (sub $.SlideSize.Width 4.0)}}" 
                           svg:height="3.506cm" 
                           svg:x="2cm" svg:y="1cm"
                           presentation:class="title">
                    <draw:text-box>
                        <text:p text:style-name="P1">{{.Title}}</text:p>
                    </draw:text-box>
                </draw:frame>
                {{end}}
                {{if .Content}}
                <draw:frame draw:style-name="gr2" draw:text-style-name="P2" draw:layer="layout"
                           svg:width="{{printf "%.2fcm" (sub $.SlideSize.Width 4.0)}}" 
                           svg:height="13.23cm" 
                           svg:x="2cm" svg:y="5.5cm"
                           presentation:class="outline">
                    <draw:text-box>
                        <text:p text:style-name="P2">{{.Content}}</text:p>
                    </draw:text-box>
                </draw:frame>
                {{end}}
                {{range .TextBoxes}}
                <draw:frame draw:style-name="gr2" draw:layer="layout"
                           svg:width="{{.Width}}" svg:height="{{.Height}}" 
                           svg:x="{{.X}}" svg:y="{{.Y}}"
                           presentation:class="outline">
                    <draw:text-box>
                        <text:p>
                            <text:span text:style-name="{{generateStyleName .Style}}">{{.Content}}</text:span>
                        </text:p>
                    </draw:text-box>
                </draw:frame>
                {{end}}
                {{range .Images}}
                <draw:frame draw:style-name="gr2" draw:layer="layout"
                           svg:width="{{.Width}}" svg:height="{{.Height}}" 
                           svg:x="{{.X}}" svg:y="{{.Y}}"
                           presentation:class="graphic">
                    <draw:image xlink:href="{{.Name}}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>
                </draw:frame>
                {{end}}
            </draw:page>
            {{end}}
        </office:presentation>
    </office:body>
</office:document-content>`

	styleCounter := 0
	tmpl, err := template.New("content").Funcs(template.FuncMap{
		"generateStyleName": func(style TextStyle) string {
			styleCounter++
			return fmt.Sprintf("T%d", styleCounter)
		},
		"sub": func(a, b float64) float64 {
			return a - b
		},
	}).Parse(contentTemplate)
	if err != nil {
		return err
	}
	return tmpl.Execute(writer, g)
}

func (g *ODPGenerator) writeStyles(writer io.Writer) error {
	stylesTemplate := `<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                       xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
                       xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
                       xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
                       xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"
                       xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
                       xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
                       xmlns:xlink="http://www.w3.org/1999/xlink">
    <office:styles>
        {{if and .Background (eq .Background.Type 0)}}
        <draw:fill-image draw:name="backgroundImage" xlink:href="{{.Background.Name}}" xlink:show="embed" xlink:actuate="onLoad"/>
        {{end}}
        {{range $index, $slide := .Slides}}
            {{if and $slide.Background (eq $slide.Background.Type 0)}}
            <draw:fill-image draw:name="slideBackground{{$index}}" xlink:href="{{$slide.Background.Name}}" xlink:show="embed" xlink:actuate="onLoad"/>
            {{end}}
        {{end}}
        {{$counter := 0}}
        {{range .Slides}}
            {{range .TextBoxes}}
                {{$counter = inc $counter}}
                <style:style style:name="T{{$counter}}" style:family="text">
                    <style:text-properties
                        {{if .Style.FontFamily}}fo:font-family="{{.Style.FontFamily}}"{{end}}
                        {{if .Style.FontSize}}fo:font-size="{{.Style.FontSize}}"{{end}}
                        {{if .Style.Color}}fo:color="{{.Style.Color}}"{{end}}
                        {{if .Style.Bold}}fo:font-weight="bold"{{end}}
                        {{if .Style.Italic}}fo:font-style="italic"{{end}}
                    />
                </style:style>
            {{end}}
        {{end}}
    </office:styles>
    <office:master-styles>
        <style:master-page style:name="Default" style:page-layout-name="PM1">
            <style:drawing-page-properties 
                presentation:background-visible="true"
                presentation:background-objects-visible="true"
                draw:fill="bitmap"
                draw:fill-image-name="backgroundImage"
                style:repeat="stretch"
                draw:background-size="border"/>
        </style:master-page>
    </office:master-styles>
    <office:automatic-styles>
        <style:page-layout style:name="PM1">
            <style:page-layout-properties fo:margin-top="0cm"
                                        fo:margin-bottom="0cm"
                                        fo:margin-left="0cm"
                                        fo:margin-right="0cm"
                                        style:print-orientation="landscape"
                                        fo:page-width="{{.SlideSize.Width}}cm"
                                        fo:page-height="{{.SlideSize.Height}}cm"/>
        </style:page-layout>
    </office:automatic-styles>
</office:document-styles>`

	tmpl, err := template.New("styles").Funcs(template.FuncMap{
		"inc": func(i int) int {
			return i + 1
		},
	}).Parse(stylesTemplate)
	if err != nil {
		return err
	}
	return tmpl.Execute(writer, g)
}

func (g *ODPGenerator) writeSettings(writer io.Writer) error {
	settingsTemplate := `<?xml version="1.0" encoding="UTF-8"?>
<office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" 
                         xmlns:xlink="http://www.w3.org/1999/xlink" 
                         xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" 
                         xmlns:ooo="http://openoffice.org/2004/office" 
                         office:version="1.2">
    <office:settings>
        <config:config-item-set config:name="ooo:view-settings">
            <config:config-item config:name="VisibleAreaTop" config:type="int">0</config:config-item>
            <config:config-item config:name="VisibleAreaLeft" config:type="int">0</config:config-item>
            <config:config-item config:name="VisibleAreaWidth" config:type="int">{{printf "%.0f" (mul .SlideSize.Width 100)}}</config:config-item>
            <config:config-item config:name="VisibleAreaHeight" config:type="int">{{printf "%.0f" (mul .SlideSize.Height 100)}}</config:config-item>
            <config:config-item-map-indexed config:name="Views">
                <config:config-item-map-entry>
                    <config:config-item config:name="ViewId" config:type="string">view1</config:config-item>
                    <config:config-item config:name="GridIsVisible" config:type="boolean">false</config:config-item>
                    <config:config-item config:name="IsSnapToGrid" config:type="boolean">true</config:config-item>
                    <config:config-item config:name="IsSnapToPageMargins" config:type="boolean">true</config:config-item>
                    <config:config-item config:name="ZoomOnPage" config:type="boolean">true</config:config-item>
                    <config:config-item config:name="SelectedPage" config:type="short">0</config:config-item>
                </config:config-item-map-entry>
            </config:config-item-map-indexed>
        </config:config-item-set>
        <config:config-item-set config:name="ooo:configuration-settings">
            <config:config-item config:name="IsPrintDate" config:type="boolean">false</config:config-item>
            <config:config-item config:name="IsPrintTime" config:type="boolean">false</config:config-item>
            <config:config-item config:name="IsPrintNotes" config:type="boolean">false</config:config-item>
            <config:config-item config:name="PrintQuality" config:type="int">0</config:config-item>
            <config:config-item-map-indexed config:name="ForbiddenCharacters">
                <config:config-item-map-entry>
                    <config:config-item config:name="Language" config:type="string">es</config:config-item>
                    <config:config-item config:name="Country" config:type="string">ES</config:config-item>
                    <config:config-item config:name="Variant" config:type="string"/>
                </config:config-item-map-entry>
            </config:config-item-map-indexed>
        </config:config-item-set>
    </office:settings>
</office:document-settings>`

	tmpl, err := template.New("settings").Funcs(template.FuncMap{
		"mul": func(a, b float64) float64 {
			return a * b
		},
	}).Parse(settingsTemplate)
	if err != nil {
		return err
	}
	return tmpl.Execute(writer, g)
}

func (g *ODPGenerator) writeConfigurations(writer io.Writer) error {
	configTemplate := `<?xml version="1.0" encoding="UTF-8"?>
<oor:component-data xmlns:oor="http://openoffice.org/2001/registry" 
                    xmlns:xs="http://www.w3.org/2001/XMLSchema" 
                    oor:name="Accelerator" 
                    oor:package="org.openoffice.Office">
    <node oor:name="PresentationCommands">
        <node oor:name="Defaults">
            <node oor:name="Modules">
                <node oor:name="com.sun.star.presentation.PresentationDocument"/>
            </node>
        </node>
    </node>
</oor:component-data>`

	_, err := writer.Write([]byte(configTemplate))
	return err
}

func (g *ODPGenerator) writeManifest(writer io.Writer) error {
	manifestTemplate := `<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
    <manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.presentation" manifest:full-path="/"/>
    <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="content.xml"/>
    <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="styles.xml"/>
    <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="settings.xml"/>
    <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="configurations2/accelerator/current.xml"/>
    {{if .Background}}
    <manifest:file-entry manifest:media-type="image/{{extension .Background.Name}}" manifest:full-path="{{.Background.Name}}"/>
    {{end}}
    {{range .Slides}}
        {{if .Background}}
        <manifest:file-entry manifest:media-type="image/{{extension .Background.Name}}" manifest:full-path="{{.Background.Name}}"/>
        {{end}}
        {{range .Images}}
    <manifest:file-entry manifest:media-type="image/{{extension .Name}}" manifest:full-path="{{.Name}}"/>
        {{end}}
    {{end}}
</manifest:manifest>`

	tmpl, err := template.New("manifest").Funcs(template.FuncMap{
		"extension": func(name string) string {
			ext := filepath.Ext(name)
			return strings.TrimPrefix(ext, ".")
		},
	}).Parse(manifestTemplate)
	if err != nil {
		return err
	}
	return tmpl.Execute(writer, g)
}
