import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import shutil
import time
import unicodedata
import subprocess

def limpiar_texto(texto):
    nfkd = unicodedata.normalize('NFKD', texto)
    return "".join([c for c in nfkd if not unicodedata.combining(c)]).lower().replace(" ", "_")

def obtener_ruta_git():
    try:
        subprocess.run(["git", "--version"], capture_output=True, text=True)
        return "git"
    except FileNotFoundError:
        pass

    usuario = os.getenv("USERNAME", "")
    rutas_comunes = [
        r"C:\Program Files\Git\cmd\git.exe",
        r"C:\Program Files (x86)\Git\cmd\git.exe",
        f"C:\\Users\\{usuario}\\AppData\\Local\\Programs\\Git\\cmd\\git.exe"
    ]

    for ruta in rutas_comunes:
        if os.path.exists(ruta):
            return ruta

    return "git"


class GestorCatalogoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gestor de Catálogo - RelojesOscar")
        self.root.geometry("480x560")
        self.root.config(bg="#f8fafc")

        self.ruta_imagen_original = ""
        self.git_cmd = obtener_ruta_git()

        tk.Label(root, text="📦 Nuevo Producto", font=("Inter", 16, "bold"), bg="#f8fafc", fg="#1e293b").pack(pady=20)

        form_frame = tk.Frame(root, bg="#f8fafc")
        form_frame.pack(padx=35, fill="x", expand=True)

        tk.Label(form_frame, text="Marca y Modelo:", font=("Inter", 10, "bold"), bg="#f8fafc", fg="#475569", anchor="w").pack(fill="x", pady=(0, 2))
        self.entry_nombre = tk.Entry(form_frame, font=("Inter", 11), relief="solid", bd=1)
        self.entry_nombre.pack(fill="x", pady=(0, 15), ipady=5)

        tk.Label(form_frame, text="Precio (ej: 45000 o 'Consultar'):", font=("Inter", 10, "bold"), bg="#f8fafc", fg="#475569", anchor="w").pack(fill="x", pady=(0, 2))
        self.entry_precio = tk.Entry(form_frame, font=("Inter", 11), relief="solid", bd=1)
        self.entry_precio.pack(fill="x", pady=(0, 15), ipady=5)

        tk.Label(form_frame, text="Descripción / Detalles:", font=("Inter", 10, "bold"), bg="#f8fafc", fg="#475569", anchor="w").pack(fill="x", pady=(0, 2))
        self.entry_descripcion = tk.Entry(form_frame, font=("Inter", 11), relief="solid", bd=1)
        self.entry_descripcion.pack(fill="x", pady=(0, 15), ipady=5)

        tk.Label(form_frame, text="Fotografía:", font=("Inter", 10, "bold"), bg="#f8fafc", fg="#475569", anchor="w").pack(fill="x", pady=(0, 2))

        btn_foto = tk.Button(form_frame, text="📂 Seleccionar Imagen", command=self.seleccionar_imagen, bg="#e2e8f0", fg="#1e293b", font=("Inter", 10, "bold"), relief="flat", cursor="hand2")
        btn_foto.pack(fill="x", pady=(0, 5), ipady=5)

        self.lbl_foto = tk.Label(form_frame, text="Ninguna imagen seleccionada", font=("Inter", 9, "italic"), bg="#f8fafc", fg="#64748b", anchor="w")
        self.lbl_foto.pack(fill="x", pady=(0, 10))

        btn_guardar = tk.Button(root, text="🚀 Guardar y Publicar en la Web", command=self.procesar_producto, bg="#10b981", fg="white", font=("Inter", 11, "bold"), relief="flat", cursor="hand2")
        btn_guardar.pack(padx=35, fill="x", pady=20, ipady=10)

    def seleccionar_imagen(self):
        archivo = filedialog.askopenfilename(title="Seleccionar imagen", filetypes=[("Imágenes", "*.jpg *.jpeg *.png *.webp")])
        if archivo:
            self.ruta_imagen_original = archivo
            nombre_archivo = os.path.basename(archivo)
            self.lbl_foto.config(text=f"Seleccionada: {nombre_archivo}", fg="#059669")

    def sincronizar_antes_de_editar(self):
        """
        Paso clave: trae los últimos cambios del remoto ANTES de tocar
        el Excel o copiar imágenes. Así evitamos que tu copia local del
        .xlsx quede desactualizada respecto a GitHub, que es la causa
        más común del error 'non-fast-forward' con archivos binarios.
        """
        subprocess.run([self.git_cmd, "config", "user.name", "RelojesOscar"], capture_output=True)
        subprocess.run([self.git_cmd, "config", "user.email", "catalogo@relojesoscar.com"], capture_output=True)

        pull = subprocess.run(
            [self.git_cmd, "pull", "origin", "main", "--allow-unrelated-histories", "--no-rebase"],
            capture_output=True, text=True
        )

        if pull.returncode != 0:
            # Si hubo conflicto real (por ejemplo el .xlsx cambió en ambos lados
            # de formas incompatibles), avisamos y cancelamos en lugar de
            # seguir y arriesgar corromper el catálogo.
            if "CONFLICT" in pull.stdout or "CONFLICT" in pull.stderr:
                subprocess.run([self.git_cmd, "merge", "--abort"], capture_output=True)
                messagebox.showerror(
                    "Conflicto al sincronizar",
                    "Hay una versión del catálogo en GitHub que choca con la tuya "
                    "(probablemente alguien más subió un producto).\n\n"
                    "Se cancelaron los cambios locales para no corromper el archivo.\n"
                    "Pedile a un técnico que revise el repositorio, o volvé a intentar "
                    "en unos minutos."
                )
                return False
            else:
                messagebox.showwarning(
                    "Aviso de sincronización",
                    "No se pudo traer los últimos cambios de GitHub antes de guardar:\n"
                    f"{pull.stderr}\n\nSe intentará continuar igual."
                )
        return True

    def procesar_producto(self):
        nombre = self.entry_nombre.get().strip()
        precio = self.entry_precio.get().strip()
        descripcion = self.entry_descripcion.get().strip()

        if not nombre or not precio or not self.ruta_imagen_original:
            messagebox.showerror("Campos incompletos", "Por favor completa la Marca/Modelo, el Precio y selecciona una imagen.")
            return

        if not os.path.exists(".git"):
            messagebox.showerror(
                "Carpeta incorrecta",
                "Esta carpeta NO es un repositorio de Git.\n\n"
                "Asegúrate de ejecutar este script dentro de la carpeta donde tienes tu proyecto web."
            )
            return

        excel_path = "Descripcion.xlsx"

        if os.path.exists(excel_path):
            try:
                os.rename(excel_path, excel_path)
            except PermissionError:
                messagebox.showerror("Archivo bloqueado", "El archivo 'Descripcion.xlsx' está abierto en Excel.\n\nPor favor ciérralo y vuelve a intentarlo.")
                return

        # 1) Sincronizar PRIMERO, antes de generar cualquier archivo nuevo
        if not self.sincronizar_antes_de_editar():
            return

        os.makedirs("img", exist_ok=True)

        ext = os.path.splitext(self.ruta_imagen_original)[1]
        nombre_limpio = limpiar_texto(nombre)
        nuevo_nombre_foto = f"{nombre_limpio}_{int(time.time())}{ext}"
        destino_foto = os.path.join("img", nuevo_nombre_foto)

        try:
            shutil.copy(self.ruta_imagen_original, destino_foto)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo copiar la imagen:\n{e}")
            return

        try:
            if os.path.exists(excel_path):
                df = pd.read_excel(excel_path)
            else:
                df = pd.DataFrame(columns=["nombre", "precio", "descripcion", "imagen"])

            try:
                if "." in precio or "," in precio:
                    precio_val = float(precio.replace(",", "."))
                else:
                    precio_val = int(precio)
            except ValueError:
                precio_val = precio

            nuevo_registro = pd.DataFrame([{
                "nombre": nombre,
                "precio": precio_val,
                "descripcion": descripcion,
                "imagen": nuevo_nombre_foto
            }])

            df = pd.concat([df, nuevo_registro], ignore_index=True)
            df.to_excel(excel_path, index=False)

        except Exception as e:
            messagebox.showerror("Error de Excel", f"No se pudo actualizar el archivo Excel:\n{e}")
            return

        # 2) Commit y push, ya con el remoto sincronizado de antemano
        try:
            subprocess.run([self.git_cmd, "add", "."], check=True)

            commit = subprocess.run(
                [self.git_cmd, "commit", "-m", f"Agregado producto: {nombre}"],
                capture_output=True, text=True
            )
            if commit.returncode != 0 and "nothing to commit" not in commit.stdout.lower():
                messagebox.showwarning("Aviso", f"No se pudo confirmar el cambio:\n{commit.stderr}")
                return

            push = subprocess.run([self.git_cmd, "push", "-u", "origin", "main"], capture_output=True, text=True)

            if push.returncode == 0:
                messagebox.showinfo("¡Éxito total!", f"¡El producto '{nombre}' se guardó y se publicó en la web correctamente!")
                self.entry_nombre.delete(0, tk.END)
                self.entry_precio.delete(0, tk.END)
                self.entry_descripcion.delete(0, tk.END)
                self.lbl_foto.config(text="Ninguna imagen seleccionada", fg="#64748b")
                self.ruta_imagen_original = ""
            else:
                # Si aun así falló (ej. alguien subió algo justo en el medio),
                # reintentamos una vez más con pull + push antes de rendirnos.
                retry_pull = subprocess.run(
                    [self.git_cmd, "pull", "origin", "main", "--no-rebase"],
                    capture_output=True, text=True
                )
                if retry_pull.returncode == 0:
                    retry_push = subprocess.run([self.git_cmd, "push", "-u", "origin", "main"], capture_output=True, text=True)
                    if retry_push.returncode == 0:
                        messagebox.showinfo("¡Éxito total!", f"¡El producto '{nombre}' se guardó y se publicó en la web correctamente!")
                        self.entry_nombre.delete(0, tk.END)
                        self.entry_precio.delete(0, tk.END)
                        self.entry_descripcion.delete(0, tk.END)
                        self.lbl_foto.config(text="Ninguna imagen seleccionada", fg="#64748b")
                        self.ruta_imagen_original = ""
                        return

                messagebox.showwarning(
                    "Aviso parcial",
                    "Se guardó localmente, pero hubo un problema al sincronizar con GitHub:\n" + push.stderr
                )
        except Exception as ex:
            messagebox.showwarning("Aviso", f"Se guardaron los archivos localmente, pero falló Git: {ex}")


if __name__ == "__main__":
    root = tk.Tk()
    app = GestorCatalogoApp(root)
    root.mainloop()
