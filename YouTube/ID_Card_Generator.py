# Importing the modules
from PyQt5 import QtCore, QtGui, QtWidgets
from PIL import Image, ImageDraw, ImageFont
import random
import os
import qrcode
import openpyxl
import sys

# Creating the Ui_Form class
class Ui_Form(object):
    # Defining the setupUi method
    def setupUi(self, Form):
        # Setting the object name, size, font and style sheet of the form
        Form.setObjectName("Form")
        Form.resize(899, 694)
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        Form.setFont(font)
        Form.setStyleSheet("QWidget{\n"
"background:rgb(85, 170, 255);\n"
"\n"
"}")
        # Creating the push button and connecting it to the capture method
        self.pushButton = QtWidgets.QPushButton(Form)
        self.pushButton.setGeometry(QtCore.QRect(320, 550, 251, 61))
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton.setFont(font)
        self.pushButton.clicked.connect(self.capture)
        self.pushButton.setStyleSheet("QPushButton{\n"
"border:3px solid black;\n"
"border-radius:15px;\n"
"background:blue;\n"
"color:white;\n"
"}\n"
"\n"
"QPushButton:hover{\n"
"border:1px solid gray;\n"
"border-radius:15px;\n"
"background:black;\n"
"color:white;\n"
"}")
        self.pushButton.setObjectName("pushButton")
        # Creating the clear button and connecting it to the clear method
        self.clearButton = QtWidgets.QPushButton(Form)
        self.clearButton.setGeometry(QtCore.QRect(650, 550, 151, 61))
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.clearButton.setFont(font)
        self.clearButton.clicked.connect(self.clear)
        self.clearButton.setStyleSheet("QPushButton{\n"
"border:3px solid black;\n"
"border-radius:15px;\n"
"background:blue;\n"
"color:white;\n"
"}\n"
"\n"
"QPushButton:hover{\n"
"border:1px solid gray;\n"
"border-radius:15px;\n"
"background:black;\n"
"color:white;\n"
"}")
        self.clearButton.setObjectName("clearButton")
        
        # Creating the label for the title
        self.label = QtWidgets.QLabel(Form)
        self.label.setGeometry(QtCore.QRect(350, 30, 351, 61))
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(95)
        font.setPixelSize(30)
        self.label.setFont(font)
        self.label.setObjectName("label")

        # Creating the label for the name
        self.label_2 = QtWidgets.QLabel(Form)
        self.label_2.setGeometry(QtCore.QRect(70, 150, 201, 21))
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        # Creating the label for the email
        self.label_3 = QtWidgets.QLabel(Form)
        self.label_3.setGeometry(QtCore.QRect(70, 230, 311, 21))
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")
        # Creating the label for the phone
        self.label_4 = QtWidgets.QLabel(Form)
        self.label_4.setGeometry(QtCore.QRect(70, 310, 311, 21))
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")
        # Creating the label for the address
        self.label_5 = QtWidgets.QLabel(Form)
        self.label_5.setGeometry(QtCore.QRect(70, 390, 311, 21))
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.label_5.setFont(font)
        self.label_5.setObjectName("label_5")
        # Creating the label for the QR code
        self.label_6 = QtWidgets.QLabel(Form)
        self.label_6.setGeometry(QtCore.QRect(70, 490, 361, 61))
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.label_6.setFont(font)
        self.label_6.setObjectName("label_6")
        # Creating the line edit for the name
        self.lineEdit = QtWidgets.QLineEdit(Form)
        self.lineEdit.setGeometry(QtCore.QRect(360, 140, 381, 31))
        self.lineEdit.setStyleSheet("QLineEdit{\n"
"\n"
"background:white;\n"
"}")
        self.lineEdit.setObjectName("lineEdit")
        # Creating the line edit for the email
        self.lineEdit_2 = QtWidgets.QLineEdit(Form)
        self.lineEdit_2.setGeometry(QtCore.QRect(360, 220, 381, 31))
        self.lineEdit_2.setStyleSheet("QLineEdit{\n"
"\n"
"background:white;\n"
"}")
        self.lineEdit_2.setObjectName("lineEdit_2")
        # Creating the line edit for the phone
        self.lineEdit_3 = QtWidgets.QLineEdit(Form)
        self.lineEdit_3.setGeometry(QtCore.QRect(360, 300, 381, 31))
        self.lineEdit_3.setStyleSheet("QLineEdit{\n"
"\n"
"background:white;\n"
"}")
        self.lineEdit_3.setObjectName("lineEdit_3")
        # Creating the line edit for the address
        self.lineEdit_4 = QtWidgets.QLineEdit(Form)
        self.lineEdit_4.setGeometry(QtCore.QRect(360, 390, 381, 31))
        self.lineEdit_4.setStyleSheet("QLineEdit{\n"
"\n"
"background:white;\n"
"}")
        self.lineEdit_4.setObjectName("lineEdit_4")
        # Setting the text for the labels and the button
        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    # Defining the retranslateUi method
    def retranslateUi(self, Form):
        # Setting the window title and the text for the labels and the button
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        self.label.setText(_translate("Form", "ID Card Generator"))
        self.label_2.setText(_translate("Form", "Enter your name:"))
        self.label_3.setText(_translate("Form", "Enter your email:"))
        self.label_4.setText(_translate("Form", "Enter your phone:"))
        self.label_5.setText(_translate("Form", "Enter your address:"))
        self.label_6.setText(_translate("Form", "Your QR code will appear here:"))
        self.pushButton.setText(_translate("Form", "Generate ID Card"))
        self.clearButton.setText(_translate("Form", "Clear"))

    # Defining the save_to_excel method
    def save_to_excel(self, name, email, phone, address):
        try:
            workbook = openpyxl.load_workbook('id.xlsx')
            sheet = workbook.active
        except FileNotFoundError:
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.append(["Name", "Email", "Phone", "Address"])

        sheet.append([name, email, phone, address])
        workbook.save('id.xlsx')

    # Defining the capture method
    def capture(self):
        # Getting the input from the line edits
        name = self.lineEdit.text()
        email = self.lineEdit_2.text()
        phone = self.lineEdit_3.text()
        address = self.lineEdit_4.text()
        # Checking if the input is valid
        if name and email and phone and address:
            # Saving the data to the Excel sheet
            self.save_to_excel(name, email, phone, address)
            # Generating a random ID number
            id_number = random.randint(100000, 999999)
            # Generating a QR code with the input data
            qr_data = f"Name: {name}\nEmail: {email}\nPhone: {phone}\nAddress: {address}\nID: {id_number}"
            qr = qrcode.QRCode(version=1, box_size=10, border=5)
            qr.add_data(qr_data)
            qr.make(fit=True)
            qr_img = qr.make_image(fill="black", back_color="white")
            # Saving the QR code image as a PNG file
            qr_path = os.path.join(os.getcwd(), "qr.png")
            qr_img.save(qr_path)
            # Creating a blank ID card image with PIL
            card_width = 800
            card_height = 500
            card = Image.new("RGB", (card_width, card_height), (255, 255, 255))
            draw = ImageDraw.Draw(card)
            # Drawing a blue rectangle on the top of the card
            draw.rectangle((0, 0, card_width, 100), fill=(85, 170, 255))
            # Drawing a black line below the rectangle
            draw.line((0, 100, card_width, 100), fill=(0, 0, 0), width=3)
            # Loading the fonts for the text
            title_font = ImageFont.truetype("arial.ttf", 40)
            text_font = ImageFont.truetype("arial.ttf", 30)
            # Writing the title on the card
            title = "ID Card"
            title_w, title_h = draw.textsize(title, font=title_font)
            title_x = (card_width - title_w) // 2
            title_y = (100 - title_h) // 2
            draw.text((title_x, title_y), title, font=title_font, )
            # Writing the name on the card
            name_label = "Name:"
            name_label_w, name_label_h = draw.textsize(name_label, font=text_font)
            name_label_x = 50
            name_label_y = 150
            draw.text((name_label_x, name_label_y), name_label, font=text_font, align="left",fill=(0, 0, 0))
            draw.text((100 + name_label_x, name_label_y), name, font=text_font,fill=(10, 1, 0))
            # Writing the email on the card
            email_label = "Email:"
            email_label_x = 50
            email_label_y = 230
            draw.text((email_label_x, email_label_y), email_label, font=text_font, align="left",fill=(0, 0, 0))
            draw.text((100 + email_label_x, email_label_y), email, font=text_font,fill=(10, 1, 0))
            # Writing the phone on the card
            phone_label = "Phone:"
            phone_label_x = 50
            phone_label_y = 310
            draw.text((phone_label_x, phone_label_y), phone_label, font=text_font, align="left",fill=(0, 0, 0))
            draw.text((110 + phone_label_x, phone_label_y), phone, font=text_font,fill=(10, 1, 0))
            # Writing the address on the card
            address_label = "Address:"
            address_label_x = 50
            address_label_y = 390
            draw.text((address_label_x, address_label_y), address_label, font=text_font, align="left",fill=(0, 0, 0))
            draw.text((130 + address_label_x, address_label_y), address, font=text_font,fill=(10, 1, 0))
            # Drawing the QR code on the card
            qr_image = Image.open(qr_path)
            qr_image = qr_image.resize((150, 150))
            card.paste(qr_image, (550, 300))
            # Displaying the ID card in a new window
            card.show()
            # Deleting the temporary QR code image file
            os.remove(qr_path)

    # Defining the clear method
    def clear(self):
        self.lineEdit.clear()
        self.lineEdit_2.clear()
        self.lineEdit_3.clear()
        self.lineEdit_4.clear()

if __name__ == "__main__":
    # Creating the Qt application
    app = QtWidgets.QApplication(sys.argv)
    # Creating the form
    Form = QtWidgets.QWidget()
    # Creating the UI object
    ui = Ui_Form()
    # Setting up the UI
    ui.setupUi(Form)
    # Showing the form
    Form.show()
    # Executing the application
    sys.exit(app.exec_())
