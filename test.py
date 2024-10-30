import qrcode
import os
import customtkinter as ctk  # Import customtkinter
from PIL import Image, ImageTk
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage

class QRCodeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("QR Code Generator")
        self.root.geometry("400x500")
         
        # Create an Excel workbook
        self.excel_path = 'qr_codes_with_images.xlsx'
        self.workbook = Workbook()
        self.sheet = self.workbook.active
        self.sheet.title = "QR Codes"
        self.sheet.append(["QR Code Content", "QR Code Image"])  # Header row

        # Label and Entry for data input
        self.label = ctk.CTkLabel(root, text="Enter Data for QR Code:", font=("Arial", 12))
        self.label.pack(pady=10)

        self.entry = ctk.CTkEntry(root, width=50, font=("Arial", 12))
        self.entry.pack(pady=10)

        # Button to generate QR Code
        self.generate_button = ctk.CTkButton(root, text="Generate QR Code", command=self.generate_qr, font=("Arial", 12), bg_color="green", text_color="white")
        self.generate_button.pack(pady=20)

        # Label to display the QR Code
        self.qr_code_label = ctk.CTkLabel(root,text="")  # No bg argument
        self.qr_code_label.pack()

    def generate_qr(self):
        # Get data from entry field
        data = self.entry.get()
        if not data:
            ctk.CTkMessageBox.show_error("Error", "Please enter data to generate QR Code")
            return

        # Generate QR Code
        qr = qrcode.QRCode(version=1, error_correction=qrcode.constants.ERROR_CORRECT_H, box_size=10, border=4)
        qr.add_data(data)
        qr.make(fit=True)

        # Create an image of the QR code
        qr_image = qr.make_image(fill="black", back_color="white")
        
        # Display the QR Code on the application
        qr_image_resized = qr_image.resize((200, 200), Image.LANCZOS)
        qr_photo = ImageTk.PhotoImage(qr_image_resized)
        self.qr_code_label.configure(image=qr_photo)
        self.qr_code_label.image = qr_photo

        # Temporary save QR image to add it to Excel
        temp_image_path = f"temp_{data}.png"
        qr_image.save(temp_image_path)

        # Check if the image file was created successfully
        if not os.path.exists(temp_image_path):
            ctk.CTkMessageBox.show_error("Error", f"Failed to create QR code image at '{temp_image_path}'")
            return

        # Append data and image to Excel
        self.sheet.append([data])
        img = ExcelImage(temp_image_path)
        img.width, img.height = 100, 100  # Resize for Excel
        self.sheet.add_image(img, f"B{self.sheet.max_row}")

        # Save the Excel workbook immediately after adding the QR code
        try:
            self.workbook.save(self.excel_path)
            ctk.CTkMessageBox.show_info("Success", f"QR Code generated and saved to '{self.excel_path}'")
        except Exception as e:
            ctk.CTkMessageBox.show_error("Error", f"Failed to save Excel file: {e}")
        finally:
            # Remove the temporary image file
            if os.path.exists(temp_image_path):
                os.remove(temp_image_path)

# Run the application
if __name__ == "__main__":
    ctk.set_appearance_mode("light")  # Set the appearance mode (optional)
    ctk.set_default_color_theme("blue")  # Set the color theme (optional)
    root = ctk.CTk()  # Create a CTk window
    app = QRCodeApp(root)
    root.mainloop()
