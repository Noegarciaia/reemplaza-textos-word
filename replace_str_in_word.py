from pathlib import Path
import win32com.client  # pip install pywin32
import tkinter as tk
from tkinter import simpledialog

# Crear la ventana de diálogo para ingresar los valores
root = tk.Tk()
root.withdraw()  # Oculta la ventana principal de tkinter

# Solicitar al usuario los valores de las variables
find_str = simpledialog.askstring("Entrada", "Palabra original (que está en el documento):")
replace_with = simpledialog.askstring("Entrada", "Palabra a reemplazar:")
link_find_str = simpledialog.askstring("Entrada", "Ingrese el enlace original (que esta en el documento):")
link_replace_with = simpledialog.askstring("Entrada", "Ingrese el enlace de reemplazo:")

# Configuración de rutas
current_dir = Path(__file__).parent if "__file__" in locals() else Path.cwd()
input_dir = current_dir / "input"
output_dir = current_dir / "output"
pdf_dir = output_dir / "pdf"
output_dir.mkdir(parents=True, exist_ok=True)
pdf_dir.mkdir(parents=True, exist_ok=True)

wd_replace = 2  # 2=replace all occurrences
wd_find_wrap = 1  # 1=continue search

# Iniciar Word
word_app = win32com.client.DispatchEx("Word.Application")
word_app.Visible = False
word_app.DisplayAlerts = False

# Iterar sobre cada archivo en la carpeta de entrada
for doc_file in Path(input_dir).rglob("*.doc*"):
    try:
        # Abrir el documento
        doc = word_app.Documents.Open(str(doc_file))

        # Reemplazar texto en el documento
        word_app.Selection.Find.Execute(
            FindText=find_str,
            ReplaceWith=replace_with,
            Replace=wd_replace,
            Forward=True,
            MatchCase=True,
            MatchWholeWord=False,
            MatchWildcards=True,
            MatchSoundsLike=False,
            MatchAllWordForms=False,
            Wrap=wd_find_wrap,
            Format=True,
        )

        # Reemplazar texto en formas (si las hay)
        if find_str and replace_with:
            for i in range(doc.Shapes.Count):
                if doc.Shapes(i + 1).TextFrame.HasText:
                    words = doc.Shapes(i + 1).TextFrame.TextRange.Words
                    for j in range(words.Count):
                        if words.Item(j + 1).Text == find_str:
                            words.Item(j + 1).Text = replace_with

        # Reemplazar enlaces en el documento
        if link_find_str and link_replace_with:
            for hyperlink in doc.Hyperlinks:
                if link_find_str in hyperlink.Address:
                    hyperlink.Address = hyperlink.Address.replace(link_find_str, link_replace_with)
                if link_find_str in hyperlink.TextToDisplay:
                    hyperlink.TextToDisplay = hyperlink.TextToDisplay.replace(link_find_str, link_replace_with)

        # Guardar el archivo .docx modificado en la carpeta de salida
        output_path = output_dir / f"{doc_file.stem}{doc_file.suffix}"
        doc.SaveAs(str(output_path))

        # Guardar el archivo como PDF en la carpeta de PDFs
        pdf_output_path = pdf_dir / f"{doc_file.stem}.pdf"
        doc.SaveAs(str(pdf_output_path), FileFormat=17)  # 17 es el formato PDF en Word

    except Exception as e:
        print(f"Error al procesar {doc_file.name}: {e}")
    
    finally:
        # Cerrar el documento sin guardar más cambios
        doc.Close(SaveChanges=False)

# Cerrar la aplicación de Word
word_app.Quit()

print("Procesamiento de archivos completado.")
