import qrcode
from tkinter import Tk, Label, Entry, Button, messagebox
from PIL import Image, ImageTk

class QRCodeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("QR Code Generator")
        self.root.geometry("400x400")
        self.root.config(bg="lightblue")

        # Label and Entry for data input
        self.label = Label(root, text="Enter Data for QR Code:", bg="lightblue", font=("Arial", 12))
        self.label.pack(pady=10)

        self.entry = Entry(root, width=50, font=("Arial", 12))
        self.entry.pack(pady=10)

        # Button to generate QR Code
        self.generate_button = Button(root, text="Generate QR Code", command=self.generate_qr, font=("Arial", 12), bg="green", fg="white")
        self.generate_button.pack(pady=20)

        # Label to display the QR Code
        self.qr_code_label = Label(root, bg="lightblue")
        self.qr_code_label.pack()

    def generate_qr(self):
        # Get data from entry field
        data = self.entry.get()

        if not data:
            messagebox.showerror("Error", "Please enter data to generate QR Code")
            return

        # Generate QR Code
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_H,
            box_size=10,
            border=4,
        )
        qr.add_data(data)
        qr.make(fit=True)

        # Create an image of the QR code
        qr_image = qr.make_image(fill="black", back_color="white")

        # Display the QR Code on the application
        qr_image = qr_image.resize((200, 200), Image.LANCZOS)  # Updated here
        qr_photo = ImageTk.PhotoImage(qr_image)
        self.qr_code_label.config(image=qr_photo)
        self.qr_code_label.image = qr_photo

        # Save QR code as a file
        qr_image.save("generated_qr_code.png")
        messagebox.showinfo("Success", "QR Code generated and saved as 'generated_qr_code.png'")

# Run the application
if __name__ == "__main__":
    root = Tk()
    app = QRCodeApp(root)
    root.mainloop()
