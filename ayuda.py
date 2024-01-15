import tkinter as tk

def obtener_texto():
    texto = entry.get()
    print(f"Texto obtenido: {texto}")
    # Puedes usar 'texto' como una variable en el resto del programa

root = tk.Tk()

# Crear una caja de texto
entry = tk.Entry(root)
entry.pack()

# Crear un botón para obtener el texto
boton = tk.Button(root, text="Obtener Texto", command=obtener_texto)
boton.pack()

root.mainloop()