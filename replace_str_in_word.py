from pathlib import Path
import win32com.client  # pip install pywin32
import tkinter as tk
from tkinter import simpledialog, messagebox


class AppMenu:
    def __init__(self, root):
        self.root = root
        self.selection = None
        self.setup_ui()

    def setup_ui(self):
        self.root.title("Menú de Aplicación")
        self.root.geometry("400x200")

        tk.Label(self.root, text="Seleccione una opción:").pack(pady=20)
        tk.Button(self.root, text="Buscar y reemplazar palabras", command=self.select_words).pack(pady=10)
        tk.Button(self.root, text="Buscar y reemplazar enlaces", command=self.select_links).pack(pady=10)

    def select_words(self):
        self.selection = "words"
        self.root.destroy()

    def select_links(self):
        self.selection = "links"
        self.root.destroy()


class LinkCollectorApp:
    def __init__(self, root):
        self.root = root
        self.links = []
        self.replace_with_link = ""
        self.setup_ui()

    def setup_ui(self):
        self.root.title("Reemplazo de enlaces")
        self.root.geometry("400x400")

        # Enlaces a buscar
        tk.Label(self.root, text="Enlaces a buscar:").pack(pady=5)
        self.link_listbox = tk.Listbox(self.root, height=10)
        self.link_listbox.pack(fill=tk.BOTH, padx=10, pady=5)

        # Campo para agregar enlace
        self.link_entry = tk.Entry(self.root)
        self.link_entry.pack(fill=tk.X, padx=10, pady=5)

        # Botón para agregar enlace
        tk.Button(self.root, text="+ Agregar", command=self.add_link).pack(pady=5)

        # Campo para ingresar el enlace de reemplazo
        tk.Label(self.root, text="Enlace de reemplazo:").pack(pady=10)
        self.replace_entry = tk.Entry(self.root)
        self.replace_entry.pack(fill=tk.X, padx=10, pady=5)

        # Botón para confirmar
        tk.Button(self.root, text="Continuar", command=self.confirm_links).pack(pady=10)

    def add_link(self):
        link = self.link_entry.get().strip()
        if link:
            self.links.append(link)
            self.link_listbox.insert(tk.END, link)
            self.link_entry.delete(0, tk.END)

    def confirm_links(self):
        self.replace_with_link = self.replace_entry.get().strip()
        if not self.links or not self.replace_with_link:
            messagebox.showerror("Error", "Debe agregar al menos un enlace y definir el enlace de reemplazo.")
        else:
            self.root.destroy()


class ProgressDialog:
    def __init__(self, root, total_files):
        self.root = root
        self.total_files = total_files
        self.current_file = 0
        self.setup_ui()

    def setup_ui(self):
        self.root.title("Progreso")
        self.root.geometry("400x150")
        self.label = tk.Label(self.root, text="Procesando archivos...")
        self.label.pack(pady=10)
        self.progress = tk.Label(self.root, text=f"0/{self.total_files} archivos procesados")
        self.progress.pack(pady=10)
        self.current_file_label = tk.Label(self.root, text="")
        self.current_file_label.pack(pady=10)

    def update_progress(self, current_file, file_name):
        self.current_file += 1
        self.progress.config(text=f"{self.current_file}/{self.total_files} archivos procesados")
        self.current_file_label.config(text=f"Procesando: {file_name}")
        self.root.update()


# Mostrar el menú principal
root = tk.Tk()
menu = AppMenu(root)
root.mainloop()

if menu.selection == "words":
    find_str = simpledialog.askstring("Entrada", "Palabra original (que está en el documento):")
    replace_with_word = simpledialog.askstring("Entrada", "Palabra a reemplazar:")
    if not find_str or not replace_with_word:
        print("Debe ingresar la palabra original y la de reemplazo. Saliendo...")
        exit(1)

elif menu.selection == "links":
    root = tk.Tk()
    app = LinkCollectorApp(root)
    root.mainloop()
    if not app.links or not app.replace_with_link:
        print("Debe ingresar los enlaces y el enlace de reemplazo. Saliendo...")
        exit(1)

# Configuración de rutas
current_dir = Path(__file__).parent if "__file__" in locals() else Path.cwd()
input_dir = current_dir / "input"
output_dir = current_dir / "output"
pdf_dir = output_dir / "pdf"
output_dir.mkdir(parents=True, exist_ok=True)
pdf_dir.mkdir(parents=True, exist_ok=True)

# Parámetros de Word
wd_replace = 2
wd_find_wrap = 1

# Obtener la lista de archivos
doc_files = list(Path(input_dir).rglob("*.doc*"))

# Crear diálogo de progreso
root = tk.Tk()
progress_dialog = ProgressDialog(root, len(doc_files))

# Iniciar Word
word_app = win32com.client.DispatchEx("Word.Application")
word_app.Visible = False
word_app.DisplayAlerts = False

# Procesar los archivos
for doc_file in doc_files:
    try:
        progress_dialog.update_progress(progress_dialog.current_file, doc_file.name)
        doc = word_app.Documents.Open(str(doc_file))

        if menu.selection == "words":
            word_app.Selection.Find.Execute(
                FindText=find_str,
                ReplaceWith=replace_with_word,
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
            for i in range(1, doc.Shapes.Count + 1):
                shape = doc.Shapes(i)
                if shape.Type == 17 and shape.TextFrame.HasText:
                    text_range = shape.TextFrame.TextRange
                    text_range.Find.Execute(FindText=find_str, ReplaceWith=replace_with_word, Replace=wd_replace)

        elif menu.selection == "links":
            for link_to_find in app.links:
                for hyperlink in doc.Hyperlinks:
                    if link_to_find in (hyperlink.Address or ""):
                        hyperlink.Address = app.replace_with_link
                    if link_to_find in (hyperlink.TextToDisplay or ""):
                        hyperlink.TextToDisplay = app.replace_with_link

        output_path = output_dir / f"{doc_file.stem}{doc_file.suffix}"
        doc.SaveAs(str(output_path))
        pdf_output_path = pdf_dir / f"{doc_file.stem}.pdf"
        doc.SaveAs(str(pdf_output_path), FileFormat=17)

    except Exception as e:
        print(f"Error al procesar {doc_file.name}: {e}")

    finally:
        doc.Close(SaveChanges=False)

word_app.Quit()
root.destroy()
print("Procesamiento de archivos completado.")
