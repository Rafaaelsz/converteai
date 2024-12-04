import os
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
from tkinter import PhotoImage

# Definições de Cores e Fontes
BACKGROUND_COLOR = "#F0F4F8"  # Azul claro suave
BUTTON_COLOR = "#6C757D"      # Cinza escuro para botões
BUTTON_HOVER_COLOR = "#5A6268"  # Cinza mais escuro no hover
TEXT_COLOR = "#212529"         # Texto em cinza escuro
TITLE_COLOR = "#007BFF"        # Cor de título azul vibrante
FONT = ("Helvetica", 12)
BUTTON_FONT = ("Helvetica", 10, "bold")

# Variável global para armazenar o diretório de saída
output_directory = None

# Função para selecionar o diretório de saída
def select_output_directory():
    global output_directory
    output_directory = filedialog.askdirectory()
    if output_directory:
        messagebox.showinfo("Diretório Selecionado", f"Os arquivos serão salvos em:\n{output_directory}")
    else:
        messagebox.showwarning("Atenção", "Nenhum diretório foi selecionado.")

# Função para salvar o arquivo no diretório selecionado
def save_file_in_directory(output_filename):
    if not output_directory:
        messagebox.showerror("Erro", "Nenhum diretório foi selecionado. Por favor, selecione um diretório de saída primeiro.")
        return None
    output_path = os.path.join(output_directory, output_filename)
    return output_path

# Função para abrir o diretório de saída ao final
def open_directory(directory):
    if os.name == 'nt':  # Windows
        os.startfile(directory)
    elif os.name == 'posix':  # macOS ou Linux
        os.system(f'open "{directory}"' if os.sys.platform == 'darwin' else f'xdg-open "{directory}"')

# Função para carregamento tardio de bibliotecas
def import_libraries():
    global Image, Converter, register_heif_opener
    from PIL import Image
    from pdf2docx import Converter
    from pillow_heif import register_heif_opener

# Funções de conversão
def jpeg_to_pdf(image_paths):
    import_libraries()
    try:
        for image_path in image_paths:
            image = Image.open(image_path)
            if image.mode != 'RGB':
                image = image.convert('RGB')
            output_filename = f"{os.path.basename(image_path).split('.')[0]}.pdf"
            output_path = save_file_in_directory(output_filename)
            if output_path:
                image.save(output_path, "PDF", resolution=100.0)
        open_directory(output_directory)
        messagebox.showinfo("Sucesso", "Todas as imagens foram convertidas com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao converter imagens: {e}")

def jpeg_to_png(image_paths):
    import_libraries()
    try:
        for image_path in image_paths:
            image = Image.open(image_path)
            output_filename = f"{os.path.basename(image_path).split('.')[0]}.png"
            output_path = save_file_in_directory(output_filename)
            if output_path:
                image.save(output_path, "PNG")
        open_directory(output_directory)
        messagebox.showinfo("Sucesso", "Todas as imagens JPEG foram convertidas para PNG!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao converter imagens JPEG: {e}")

def heic_to_jpeg(heic_paths):
    import_libraries()
    try:
        register_heif_opener()
        for heic_path in heic_paths:
            image = Image.open(heic_path)
            output_filename = f"{os.path.basename(heic_path).split('.')[0]}.jpeg"
            output_path = save_file_in_directory(output_filename)
            if output_path:
                image.save(output_path, "JPEG")
        open_directory(output_directory)
        messagebox.showinfo("Sucesso", "Todos os arquivos HEIC foram convertidos para JPEG!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao converter HEIC para JPEG: {e}")

def png_to_pdf(png_paths):
    import_libraries()
    try:
        for png_path in png_paths:
            image = Image.open(png_path)
            if image.mode != 'RGB':
                image = image.convert('RGB')
            output_filename = f"{os.path.basename(png_path).split('.')[0]}.pdf"
            output_path = save_file_in_directory(output_filename)
            if output_path:
                image.save(output_path, "PDF")
        open_directory(output_directory)
        messagebox.showinfo("Sucesso", "Todas as imagens PNG foram convertidas para PDF!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao converter imagens PNG: {e}")

def pdf_to_docx(pdf_paths):
    import_libraries()
    try:
        for pdf_path in pdf_paths:
            output_filename = f"{os.path.basename(pdf_path).split('.')[0]}.docx"
            output_path = save_file_in_directory(output_filename)
            if output_path:
                cv = Converter(pdf_path)
                cv.convert(output_path, start=0, end=None)
                cv.close()
        open_directory(output_directory)
        messagebox.showinfo("Sucesso", "Todos os PDFs foram convertidos para DOCX!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao converter PDF para DOCX: {e}")

def jfif_to_jpeg(jfif_paths):
    import_libraries()
    try:
        for jfif_path in jfif_paths:
            image = Image.open(jfif_path)
            output_filename = f"{os.path.basename(jfif_path).split('.')[0]}.jpeg"
            output_path = save_file_in_directory(output_filename)
            if output_path:
                image.save(output_path, "JPEG")
        open_directory(output_directory)
        messagebox.showinfo("Sucesso", "Todos os arquivos JFIF foram convertidos para JPEG!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao converter JFIF para JPEG: {e}")

# Função para rodar as conversões em threads
def run_conversion_in_thread(conversion_function, *args):
    thread = threading.Thread(target=conversion_function, args=args)
    thread.start()

# Funções para selecionar arquivos
def select_multiple_files(file_types, conversion_function):
    file_paths = filedialog.askopenfilenames(filetypes=file_types)
    if file_paths:
        run_conversion_in_thread(conversion_function, file_paths)

# Função para estilizar botões
def style_button(button):
    button.config(
        bg=BUTTON_COLOR,
        fg="white",
        font=BUTTON_FONT,
        relief="flat",
        activebackground=BUTTON_HOVER_COLOR,
        activeforeground="white",
        cursor="hand2",
        height=2,
        width=40
    )

# Configuração da janela
root = tk.Tk()
root.title("Conversor de Arquivos v1")
root.geometry("600x700")
root.configure(bg=BACKGROUND_COLOR)

# Título
title = tk.Label(root, text="Conversor de Arquivos", bg=BACKGROUND_COLOR, fg=TITLE_COLOR, font=("Helvetica", 18, "bold"))
title.pack(pady=20)

# Botões
button_frame = tk.Frame(root, bg=BACKGROUND_COLOR)
button_frame.pack(pady=10)

btn_select_directory = tk.Button(button_frame, text="Selecionar Diretório de Saída", command=select_output_directory)
style_button(btn_select_directory)
btn_select_directory.pack(pady=10)

btn_jpeg = tk.Button(button_frame, text="Converter JPEG para PDF", 
                     command=lambda: select_multiple_files([("JPEG files", "*.jpeg;*.jpg")], jpeg_to_pdf))
style_button(btn_jpeg)
btn_jpeg.pack(pady=10)

btn_jpeg_to_png = tk.Button(button_frame, text="Converter JPEG para PNG", 
                            command=lambda: select_multiple_files([("JPEG files", "*.jpeg;*.jpg")], jpeg_to_png))
style_button(btn_jpeg_to_png)
btn_jpeg_to_png.pack(pady=10)

btn_heic = tk.Button(button_frame, text="Converter HEIC para JPEG", 
                     command=lambda: select_multiple_files([("HEIC files", "*.heic")], heic_to_jpeg))
style_button(btn_heic)
btn_heic.pack(pady=10)

btn_png = tk.Button(button_frame, text="Converter PNG para PDF", 
                    command=lambda: select_multiple_files([("PNG files", "*.png")], png_to_pdf))
style_button(btn_png)
btn_png.pack(pady=10)

btn_pdf = tk.Button(button_frame, text="Converter PDF para DOCX", 
                    command=lambda: select_multiple_files([("PDF files", "*.pdf")], pdf_to_docx))
style_button(btn_pdf)
btn_pdf.pack(pady=10)

btn_jfif_to_jpeg = tk.Button(button_frame, text="Converter JFIF para JPEG", 
                             command=lambda: select_multiple_files([("JFIF files", "*.jfif")], jfif_to_jpeg))
style_button(btn_jfif_to_jpeg)
btn_jfif_to_jpeg.pack(pady=10)

root.mainloop()
