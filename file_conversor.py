import os
import tkinter as tk
from tkinter import filedialog, messagebox
import threading

BACKGROUND_COLOR = "#F0F4F8"  
BUTTON_COLOR = "#6C757D"     
BUTTON_HOVER_COLOR = "#5A6268" 
TEXT_COLOR = "#212529"        
TITLE_COLOR = "#007BFF"       
FONT = ("Helvetica", 12)
BUTTON_FONT = ("Helvetica", 10, "bold")

class FileConversorApp: 
    def __init__(self, root):
        self.output_directory = None
        self.root = root
        self.setup_ui()

    def setup_ui(self):
        self.root.title("Conversor de Arquivos")
        self.root.configure(bg=BACKGROUND_COLOR)

        title = tk.Label(self.root, text="Conversor de Arquivos", bg=BACKGROUND_COLOR, fg=TITLE_COLOR, font=("Helvetica", 18, "bold"))
        title.pack(pady=20)

        button_frame = tk.Frame(self.root, bg=BACKGROUND_COLOR)
        button_frame.pack(pady=10, padx=10)

        def style_button(button):
            button.config(font=BUTTON_FONT, bg=BUTTON_COLOR,fg=TEXT_COLOR, activebackground=BUTTON_HOVER_COLOR, activeforeground=TEXT_COLOR, bd=0, relief="flat", padx=10, pady=5)

        btn_select_directory = tk.Button(button_frame, text="Selecionar Diretório que o Arquivo será Salvo", command=self.select_output_directory)
        style_button(btn_select_directory)
        btn_select_directory.pack(pady=10, fill=tk.X)

        btn_jpeg = tk.Button(button_frame, text="Converter JPEG para PDF", command=lambda: self.select_multiple_files([("JPEG files", "*.jpeg;*.jpg")], self.jpeg_to_pdf))
        style_button(btn_jpeg)
        btn_jpeg.pack(pady=10, fill=tk.X)

        btn_jpeg_to_png = tk.Button(button_frame, text="Converter JPEG para PNG", command=lambda: self.select_multiple_files([("JPEG files", "*.jpeg;*.jpg")], self.jpeg_to_png))
        style_button(btn_jpeg_to_png)
        btn_jpeg_to_png.pack(pady=10, fill=tk.X)

        btn_heic = tk.Button(button_frame, text="Converter HEIC para JPEG", command=lambda: self.select_multiple_files([("HEIC files", "*.heic")], self.heic_to_jpeg))
        style_button(btn_heic)
        btn_heic.pack(pady=10, fill=tk.X)

        btn_png = tk.Button(button_frame, text="Converter PNG para PDF", command=lambda: self.select_multiple_files([("PNG files", "*.png")], self.png_to_pdf))
        style_button(btn_png)
        btn_png.pack(pady=10, fill=tk.X)

        btn_pdf = tk.Button(button_frame, text="Converter PDF para DOCX", command=lambda: self.select_multiple_files([("PDF files", "*.pdf")], self.pdf_to_docx))
        style_button(btn_pdf)
        btn_pdf.pack(pady=10, fill=tk.X)

        btn_jfif_to_jpeg = tk.Button(button_frame, text="Converter JFIF para JPEG", command=lambda: self.select_multiple_files([("JFIF files", "*.jfif")], self.jfif_to_jpeg))
        style_button(btn_jfif_to_jpeg)
        btn_jfif_to_jpeg.pack(pady=10, fill=tk.X)

        btn_ogg_to_wav = tk.Button(button_frame, text="Converter OGG para WAV", command=lambda: self.select_multiple_files([("OGG files", "*.ogg")], self.ogg_to_wav))
        style_button(btn_ogg_to_wav)
        btn_ogg_to_wav.pack(pady=10, fill=tk.X)

    def select_output_directory(self):
        self.output_directory = filedialog.askdirectory()
        if self.output_directory:
            messagebox.showinfo("Diretório Selecionado", f"Os arquivos serão salvos em:\n{self.output_directory}")
        else:
            messagebox.showwarning("Atenção", "Nenhum diretório foi selecionado.")

    def select_multiple_files(self, filetypes, conversion_function):
        files = filedialog.askopenfilenames(filetypes=filetypes)
        if files:
            threading.Thread(target=conversion_function, args=(files,)).start()

    def save_file_in_directory(self, output_filename):
        if not self.output_directory:
            self.output_directory = os.path.join(os.path.expanduser("~"), "Downloads")
        output_path = os.path.join(self.output_directory, output_filename)
    
        base, extension = os.path.splitext(output_path)
        counter = 1
        while os.path.exists(output_path):
            output_path = f"{base}_{counter}{extension}"
            counter += 1
        
        return output_path

    def jpeg_to_pdf(self, image_paths):
        try:
            from PIL import Image
            combine = messagebox.askyesno("Combinar Arquivos", "Deseja combinar todas as imagens em um único arquivo PDF?")
            if combine:
                images = []
                for image_path in image_paths:
                    image = Image.open(image_path)
                    if image.mode != 'RGB':
                        image = image.convert('RGB')
                    images.append(image)
                output_filename = self.save_file_in_directory("combined.pdf")
                if output_filename:
                    images[0].save(output_filename, save_all=True, append_images=images[1:], resolution=100.0)
            else:
                for image_path in image_paths:
                    image = Image.open(image_path)
                    if image.mode != 'RGB':
                        image = image.convert('RGB')
                    output_filename = self.save_file_in_directory(f"{os.path.basename(image_path).split('.')[0]}.pdf")
                    if output_filename:
                        image.save(output_filename, "PDF", resolution=100.0)
            messagebox.showinfo("Sucesso", "Todas as imagens foram convertidas com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao converter imagens: {e}")

    def jpeg_to_png(self, image_paths):
        try:
            from PIL import Image
            for image_path in image_paths:
                image = Image.open(image_path)
                output_filename = self.save_file_in_directory(f"{os.path.basename(image_path).split('.')[0]}.png")
                if output_filename:
                    image.save(output_filename, "PNG")
            messagebox.showinfo("Sucesso", "Todas as imagens JPEG foram convertidas para PNG!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao converter imagens: {e}")

    def heic_to_jpeg(self, image_paths):
        try:
            from PIL import Image
            from pillow_heif import register_heif_opener
            register_heif_opener()
            for image_path in image_paths:
                image = Image.open(image_path)
                output_filename = self.save_file_in_directory(f"{os.path.basename(image_path).split('.')[0]}.jpeg")
                if output_filename:
                    image.save(output_filename, "JPEG")
            messagebox.showinfo("Sucesso", "Todas as imagens HEIC foram convertidas para JPEG!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao converter imagens: {e}")

    def png_to_pdf(self, image_paths):
        try:
            from PIL import Image
            combine = messagebox.askyesno("Combinar Arquivos", "Deseja combinar todas as imagens em um único arquivo PDF?")
            if combine:
                images = []
                for image_path in image_paths:
                    image = Image.open(image_path)
                    if image.mode != 'RGB':
                        image = image.convert('RGB')
                    images.append(image)
                output_filename = self.save_file_in_directory("combined.pdf")
                if output_filename:
                    images[0].save(output_filename, save_all=True, append_images=images[1:], resolution=100.0)
            else:
                for image_path in image_paths:
                    image = Image.open(image_path)
                    if image.mode != 'RGB':
                        image = image.convert('RGB')
                    output_filename = self.save_file_in_directory(f"{os.path.basename(image_path).split('.')[0]}.pdf")
                    if output_filename:
                        image.save(output_filename, "PDF", resolution=100.0)
            messagebox.showinfo("Sucesso", "Todas as imagens foram convertidas com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao converter imagens: {e}")

    def pdf_to_docx(self, pdf_paths):
        try:
            from pdf2docx import Converter
            for pdf_path in pdf_paths:
                output_filename = self.save_file_in_directory(f"{os.path.basename(pdf_path).split('.')[0]}.docx")
                if output_filename:
                    cv = Converter(pdf_path)
                    cv.convert(output_filename, start=0, end=None)
                    cv.close()
            messagebox.showinfo("Sucesso", "Todos os arquivos PDF foram convertidos para DOCX!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao converter arquivos PDF: {e}")

    def jfif_to_jpeg(self, image_paths):
        try:
            from PIL import Image
            for image_path in image_paths:
                image = Image.open(image_path)
                output_filename = self.save_file_in_directory(f"{os.path.basename(image_path).split('.')[0]}.jpeg")
                if output_filename:
                    image.save(output_filename, "JPEG")
            messagebox.showinfo("Sucesso", "Todas as imagens JFIF foram convertidas para JPEG!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao converter imagens: {e}")

    def ogg_to_wav(self, audio_paths):
        try:
            import soundfile as sf
            for audio_path in audio_paths:
                data, samplerate = sf.read(audio_path)
                output_filename = self.save_file_in_directory(f"{os.path.basename(audio_path).split('.')[0]}.wav")
                if output_filename:
                    sf.write(output_filename, data, samplerate)
            messagebox.showinfo("Sucesso", "Todos os arquivos OGG foram convertidos para WAV!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao converter arquivos de áudio: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = FileConversorApp(root)
    root.mainloop()