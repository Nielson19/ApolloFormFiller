import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
import datetime
import re
from pathlib import Path
import fitz  # PyMuPDF
import os
import sys  # Para manejar rutas en PyInstaller

# Función para obtener la ruta correcta de los recursos
def get_resource_path(relative_path):
    """Obtiene la ruta absoluta al recurso, funciona para desarrollo y para PyInstaller."""
    try:
        # PyInstaller crea una carpeta temporal y almacena la ruta en _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# Apollo Form Filler

def main():
    document_dir = Path('.')

    # Rutas de los contratos usando get_resource_path
    contracts = {
        "Apollo Contract 2025": get_resource_path("Templates/EMPTY APOLLO 2025 - CONTRACT.pdf"),
        "Flood Contract 2025": get_resource_path("Templates/FLOOD APOLLO 2025 - FINAL CONTRACT.pdf"),
        "Tarp Contract 2025": get_resource_path("Templates/TARP APOLLO 2025.pdf"),
        "WM Contract 2025": get_resource_path("Templates/WM APOLLO 2025.pdf"),
        "Mold Contract 2025": get_resource_path("Templates/MOLD APOLLO 2025.pdf"),
    }

    # Configurar el modo de apariencia y el tema
    ctk.set_appearance_mode("System")  # Puedes usar "Light" o "Dark"
    ctk.set_default_color_theme("blue")  # Tema azul predeterminado

    app = ctk.CTk()

    # Obtener el tamaño de la pantalla
    screen_width = app.winfo_screenwidth()
    screen_height = app.winfo_screenheight()

    # Establecer el tamaño de la ventana como un porcentaje del tamaño de la pantalla
    window_width = int(screen_width * 0.8)  # 80% del ancho de la pantalla
    window_height = int(screen_height * 0.8)  # 80% del alto de la pantalla

    # Calcular la posición para centrar la ventana
    position_top = int((screen_height / 2) - (window_height / 2))
    position_right = int((screen_width / 2) - (window_width / 2))

    app.geometry(f'{window_width}x{window_height}+{position_right}+{position_top}')
    app.title("Apollo Form Filler")

    # Configurar el redimensionamiento de la ventana
    app.grid_rowconfigure(0, weight=1)
    app.grid_columnconfigure(0, weight=1)

    # Crear el contenedor principal con Canvas y Scrollbar
    main_frame = ctk.CTkFrame(app)
    main_frame.pack(fill="both", expand=True)

    canvas = tk.Canvas(main_frame)
    scrollbar = ctk.CTkScrollbar(main_frame, orientation="vertical", command=canvas.yview)

    scrollable_frame = ctk.CTkFrame(canvas)
    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # Habilitar el scroll con la rueda del mouse
    def on_mouse_wheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    # Vincular el evento del scroll del mouse al Canvas
    canvas.bind_all("<MouseWheel>", on_mouse_wheel)

    # Función para crear campos de entrada
    def inputField(labelName):
        frame = ctk.CTkFrame(scrollable_frame)
        frame.pack(fill='x', pady=5)

        label = ctk.CTkLabel(frame, text=labelName)
        label.pack(side='left', anchor='w', padx=10)

        value = tk.StringVar()
        inputSpace = ctk.CTkEntry(frame, width=350, height=40, textvariable=value)
        inputSpace.pack(side='right', padx=10, pady=2)

        return value

    labelTitle = ctk.CTkLabel(scrollable_frame, text="PDF Form Filler", font=("Arial", 16))
    labelTitle.pack()

    # Checkboxes para selección de contrato
    contract_vars = {}
    estimated_totals = {}  # Almacenar entradas de totales estimados para cada contrato

    for contract_name in contracts.keys():
        contract_vars[contract_name] = tk.BooleanVar()
        checkbox = ctk.CTkCheckBox(scrollable_frame, text=contract_name, variable=contract_vars[contract_name])
        checkbox.pack()
        estimated_totals[contract_name] = inputField(f"Estimated Total for {contract_name}")

    labelTitle2 = ctk.CTkLabel(scrollable_frame, text="Form Fields", font=("Arial", 14))
    labelTitle2.pack()

    # Campos del formulario estándar
    insureName = inputField("Insured Name")
    address = inputField("Address")
    city = inputField("City")
    state = inputField("State")
    zipcode = inputField("Zip Code")
    homePhone = inputField("Home Phone")
    cellPhone = inputField("Cell Phone")
    email = inputField("Email")
    insuranceCompany = inputField("Insurance Company")
    dateOfLoss = inputField("Date of Loss")
    PolicyNumber = inputField("Policy Number")
    ClaimNumber = inputField("Claim Number")
    ClientName = inputField("Client Name")
    ClientName2 = inputField("Client Name 2")
    ClientName3 = inputField("Client Name 3")
    SignatureDate = inputField("Signature Date")
    OnBehalf = inputField("On Behalf")

    def generateForm():
        base_form_data = {
            'Insured': insureName.get(),
            'Date_1': str(datetime.date.today()),
            'Address': address.get(),
            'City': city.get(),
            'Zip': zipcode.get(),
            'Home_Phone': homePhone.get(),
            'Cell_Phone': cellPhone.get(),
            'Email': email.get(),
            'Insurance_Company': insuranceCompany.get(),
            'Date_of_Loss': dateOfLoss.get(),
            'Policy_Num': PolicyNumber.get(),
            'Claim_Num': ClaimNumber.get(),
            'Client_Name': ClientName.get(),
            'Client_Name_2': ClientName2.get(),
            'Client_Name_3': ClientName3.get(),
            'Date_2': SignatureDate.get(),
            'On_Behalf': OnBehalf.get(),
        }

        # Obtener la ruta del escritorio del usuario actual
        desktop_dir = Path(os.path.expanduser("~")) / "Contratos Apollo"

        # Crear una carpeta en el escritorio con el nombre del cliente
        client_name = base_form_data['Insured'].strip()  # Nombre del cliente
        output_dir = desktop_dir / client_name
        output_dir.mkdir(parents=True, exist_ok=True)

        # Generar un archivo PDF para cada contrato seleccionado
        for contractName, contract_path in contracts.items():
            if contract_vars[contractName].get():  # Solo procesar contratos seleccionados
                form_data = base_form_data.copy()
                form_data['Estimated_Total'] = estimated_totals[contractName].get()

                # Verificar si el archivo PDF existe
                if not os.path.exists(contract_path):
                    messagebox.showwarning(
                        "Archivo no encontrado",
                        f"No se encontró el archivo PDF para {contractName}.\nRuta esperada: {contract_path}"
                    )
                    continue  # Saltar al siguiente contrato

                # Generar el nombre del archivo
                output_file = output_dir / f"{contractName}.pdf"

                # Llenar y guardar el PDF
                try:
                    FieldFilling(document_dir, contract_path, form_data, output_file)
                    messagebox.showinfo(
                        "Archivo Generado",
                        f"El archivo se ha guardado en:\n{output_file}"
                    )
                except Exception as e:
                    messagebox.showerror(
                        "Error",
                        f"No se pudo generar el archivo PDF para {contractName}.\nError: {e}"
                    )

    generateButton = ctk.CTkButton(scrollable_frame, text="Generate Form", command=generateForm)
    generateButton.pack()

    app.mainloop()


def format_phone_number(phone):
    digits = re.sub(r'\D', '', phone)
    if len(digits) == 10:
        return f"({digits[:3]}) {digits[3:6]}-{digits[6:]}"
    return phone


def FieldFilling(document_dir, source_file_name, form_data, output_file):
    with fitz.open(source_file_name) as doc:
        filled = False
        for page in doc:
            for field in page.widgets():
                if field.field_type == 7:
                    field_name = field.field_name
                    if field_name in form_data:
                        input_value = form_data[field_name]
                        if 'Phone' in field_name:
                            input_value = format_phone_number(input_value)
                        field.field_value = str(input_value)
                        field.update()
                        filled = True

                    if field_name == "Estimated_Total":
                        field.field_value = str(form_data['Estimated_Total'])
                        field.update()
                        filled = True

        if filled:
            doc.save(output_file)
            print(f"Archivo PDF guardado como {output_file}")
        else:
            print(f"No fields were updated for {output_file}")


if __name__ == "__main__":
    main()