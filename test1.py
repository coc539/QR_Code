import qrcode
import os
import customtkinter as ctk
from PIL import Image
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from tkinter import messagebox
from datetime import datetime

class QRCodeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("QR Code Generator")
        self.root.geometry("400x500")
        
        # Create a unique Excel workbook filename with a timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Directory to store Excel files
        self.excel_directory = "excel_files"
        os.makedirs(self.excel_directory, exist_ok=True)  # Create directory if it doesn't exist
        self.excel_path = os.path.join(self.excel_directory, f'qr_codes_with_images_{timestamp}.xlsx')
        
        # Initialize the workbook and sheet
        self.workbook = Workbook()
        self.sheet = self.workbook.active
        self.sheet.title = "QR Codes"
        
        # Add headers to the Excel sheet
        self.sheet.append(["QR Code Content", "QR Code Image"])  # Header row
        
        # Set column width and row height to accommodate QR images and text
        self.sheet.column_dimensions['A'].width = 30  # Width for QR Code Content
        self.sheet.column_dimensions['B'].width = 20  # Width for QR Code Image
        self.row_height = 80  # Fixed height for rows with QR images

        # Directory to store QR code images
        self.image_directory = "qr_code_images"
        os.makedirs(self.image_directory, exist_ok=True)  # Create directory if it doesn't exist

        # Label and Entry for data input
        self.label = ctk.CTkLabel(root, text="Enter Data for QR Code:", font=("Arial", 12))
        self.label.pack(pady=10)

        self.entry = ctk.CTkEntry(root, width=375, font=("Arial", 12))
        self.entry.pack(pady=10)

        # Button to generate QR Code
        self.generate_button = ctk.CTkButton(root, text="Generate QR Code", command=self.generate_qr, font=("Arial", 12), fg_color="green", text_color="white")
        self.generate_button.pack(pady=20)

        # Label to display the QR Code
        self.qr_code_label = ctk.CTkLabel(root, text="")  # No bg argument
        self.qr_code_label.pack()

    def generate_qr(self):
        # Get data from entry field
        data = self.entry.get()
        if not data:
            messagebox.showerror("Error", "Please enter data to generate QR Code")
            return

        # Generate QR Code
        qr = qrcode.QRCode(version=1, error_correction=qrcode.constants.ERROR_CORRECT_H, box_size=10, border=4)
        qr.add_data(data)
        qr.make(fit=True)

        # Create an image of the QR code
        qr_image = qr.make_image(fill="black", back_color="white")

        # Define a permanent path for the QR code image
        image_filename = f"{self.image_directory}/{data.replace(' ', '_')}.png"  # Replace spaces with underscores for the filename
        qr_image.save(image_filename)

        # Use CTkImage for displaying the QR code
        qr_image_pil = Image.open(image_filename)  # Open the saved QR image
        qr_photo = ctk.CTkImage(qr_image_pil, size=(200, 200))  # Create a CTkImage

        # Display the QR Code on the application
        self.qr_code_label.configure(image=qr_photo)
        self.qr_code_label.image = qr_photo

        # Append data and image to Excel
        row_number = self.sheet.max_row + 1  # Determine the next row number
        self.sheet.append([data])  # Append the data text in the first column
        self.sheet.row_dimensions[row_number].height = self.row_height  # Set the row height

        # Insert the QR image into the Excel sheet at the correct row
        img = ExcelImage(image_filename)  # Use the permanent image path
        img.width, img.height = 100, 100  # Resize for Excel
        self.sheet.add_image(img, f"B{row_number}")  # Add image to the second column of the current row

        # Save the Excel workbook immediately after adding the QR code
        try:
            self.workbook.save(self.excel_path)
            messagebox.showinfo("Success", f"QR Code generated and saved to '{self.excel_path}'")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save Excel file: {e}")

# Run the application                    
if __name__ == "__main__":
    ctk.set_appearance_mode("light")  # Set the appearance mode (optional)
    ctk.set_default_color_theme("blue")  # Set the color theme (optional)
    root = ctk.CTk()  # Create a CTk window
    app = QRCodeApp(root)
    root.mainloop()
