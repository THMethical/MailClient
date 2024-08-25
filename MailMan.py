import sys
import imaplib
import smtplib
import email
import json
from email.mime.text import MIMEText
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QWidget, QSplitter, QListWidget, QListWidgetItem,
    QTextEdit, QComboBox, QPushButton, QLabel, QLineEdit, QMessageBox, QDialog, QFormLayout, QMenu,
    QAction, QGridLayout, QFrame, QColorDialog, QListWidgetItem, QDialogButtonBox
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor


class AccountDialog(QDialog):
    def __init__(self, account_info=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Konto bearbeiten")
        self.setGeometry(300, 300, 400, 200)
        
        self.email_input = QLineEdit(self)
        self.password_input = QLineEdit(self)
        self.password_input.setEchoMode(QLineEdit.Password)
        self.imap_input = QLineEdit(self)
        self.smtp_input = QLineEdit(self)
        
        if account_info:
            self.email_input.setText(account_info.get("email", ""))
            self.password_input.setText(account_info.get("password", ""))
            self.imap_input.setText(account_info.get("imap_server", ""))
            self.smtp_input.setText(account_info.get("smtp_server", ""))

        form_layout = QFormLayout()
        form_layout.addRow("E-Mail-Adresse:", self.email_input)
        form_layout.addRow("Passwort:", self.password_input)
        form_layout.addRow("IMAP-Server:", self.imap_input)
        form_layout.addRow("SMTP-Server:", self.smtp_input)
        
        self.save_button = QPushButton("Speichern", self)
        self.save_button.clicked.connect(self.accept)
        form_layout.addWidget(self.save_button)
        
        self.setLayout(form_layout)
        
    def get_account_info(self):
        return {
            "email": self.email_input.text(),
            "password": self.password_input.text(),
            "imap_server": self.imap_input.text(),
            "smtp_server": self.smtp_input.text()
        }


class ColorSelectionDialog(QDialog):
    def __init__(self, colors, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Farbe auswählen")
        self.setGeometry(300, 300, 300, 200)
        
        layout = QVBoxLayout()
        
        self.color_list = QListWidget()
        self.color_list.setSelectionMode(QListWidget.SingleSelection)
        
        for color_name, color_value in colors.items():
            item = QListWidgetItem(color_name)
            item.setBackground(color_value)
            self.color_list.addItem(item)
        
        layout.addWidget(self.color_list)
        
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        
        layout.addWidget(button_box)
        
        self.setLayout(layout)
    
    def selected_color(self):
        selected_item = self.color_list.currentItem()
        if selected_item:
            return selected_item.background().color()
        return None


class EmailClient(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("E-Mail-Client")
        self.setGeometry(100, 100, 800, 600)

        self.accounts = self.load_accounts()
        self.current_account = None
        self.emails = []
        self.marked_emails = {}
        
        self.color_legend = {
            "Schlechter Kunde": QColor("#FF0000"),
            "Guter Kunde": QColor("#00FF00"),
            "Interessiert / Zahlung offen": QColor("#0000FF")
        }

        self.load_marked_emails()  # Laden der markierten E-Mails

        self.initUI()

    def initUI(self):
        main_layout = QVBoxLayout()
        account_layout = QHBoxLayout()
        self.email_list = QListWidget()
        self.email_content = QTextEdit()
        self.email_content.setReadOnly(True)

        self.account_combo = QComboBox()
        self.account_combo.currentIndexChanged.connect(self.load_emails)
        account_layout.addWidget(QLabel("Konto:"))
        account_layout.addWidget(self.account_combo)

        add_account_button = QPushButton("Konto hinzufügen")
        add_account_button.clicked.connect(self.add_account)
        account_layout.addWidget(add_account_button)

        remove_account_button = QPushButton("Konto entfernen")
        remove_account_button.clicked.connect(self.remove_account)
        account_layout.addWidget(remove_account_button)

        edit_account_button = QPushButton("Konto bearbeiten")
        edit_account_button.clicked.connect(self.edit_account)
        account_layout.addWidget(edit_account_button)

        main_layout.addLayout(account_layout)

        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(self.email_list)
        splitter.addWidget(self.email_content)
        main_layout.addWidget(splitter)

        response_layout = QVBoxLayout()
        self.response_text = QTextEdit()
        response_layout.addWidget(QLabel("Antwort:"))
        response_layout.addWidget(self.response_text)

        send_button = QPushButton("Senden")
        send_button.clicked.connect(self.send_email)
        response_layout.addWidget(send_button)

        main_layout.addLayout(response_layout)

        legend_layout = QGridLayout()
        legend_layout.addWidget(QLabel("Legende:"), 0, 0)

        self.add_color_legend(legend_layout, "Schlechter Kunde", "#FF0000", 1)
        self.add_color_legend(legend_layout, "Guter Kunde", "#00FF00", 2)
        self.add_color_legend(legend_layout, "Interessiert / Zahlung offen", "#0000FF", 3)

        legend_frame = QFrame()
        legend_frame.setLayout(legend_layout)
        main_layout.addWidget(legend_frame)

        central_widget = QWidget()
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)

        self.email_list.currentItemChanged.connect(self.display_email)
        self.email_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.email_list.customContextMenuRequested.connect(self.show_context_menu)

        self.populate_account_combo()

    def add_color_legend(self, layout, color_name, color_value, row):
        color_label = QLabel(color_name)
        color_label.setStyleSheet(f"background-color: {color_value}; color: white; padding: 5px;")
        layout.addWidget(color_label, row, 0)
        self.color_legend[color_name] = QColor(color_value)

    def add_account(self):
        dialog = AccountDialog(parent=self)
        if dialog.exec_() == QDialog.Accepted:
            account_info = dialog.get_account_info()
            self.accounts.append(account_info)
            self.save_accounts()
            self.populate_account_combo()

    def remove_account(self):
        if self.account_combo.currentIndex() == -1:
            QMessageBox.warning(self, "Warnung", "Kein Konto ausgewählt.")
            return

        account_index = self.account_combo.currentIndex()
        del self.accounts[account_index]
        self.save_accounts()
        self.populate_account_combo()

    def edit_account(self):
        if self.account_combo.currentIndex() == -1:
            QMessageBox.warning(self, "Warnung", "Kein Konto ausgewählt.")
            return

        account_index = self.account_combo.currentIndex()
        account_info = self.accounts[account_index]
        
        dialog = AccountDialog(account_info=account_info, parent=self)
        if dialog.exec_() == QDialog.Accepted:
            self.accounts[account_index] = dialog.get_account_info()
            self.save_accounts()
            self.populate_account_combo()

    def populate_account_combo(self):
        self.account_combo.clear()
        for account in self.accounts:
            self.account_combo.addItem(account["email"])

    def save_accounts(self):
        with open("accounts.json", "w") as f:
            json.dump(self.accounts, f)

    def load_accounts(self):
        try:
            with open("accounts.json", "r") as f:
                return json.load(f)
        except FileNotFoundError:
            return []

    def save_marked_emails(self):
        with open("marked_emails.json", "w") as f:
            # Konvertiere die Schlüssel der markierten E-Mails in Strings
            json.dump({str(k): v.name() for k, v in self.marked_emails.items()}, f)

    def load_marked_emails(self):
        try:
            with open("marked_emails.json", "r") as f:
                # Konvertiere die Farben zurück
                self.marked_emails = {
                    k: QColor(v) for k, v in json.load(f).items()
                }
        except FileNotFoundError:
            self.marked_emails = {}
        except json.JSONDecodeError:
            self.marked_emails = {}

    def load_emails(self):
        self.email_list.clear()
        self.emails.clear()

        if self.account_combo.currentIndex() == -1:
            return

        self.current_account = self.accounts[self.account_combo.currentIndex()]
        try:
            mail = imaplib.IMAP4_SSL(self.current_account["imap_server"])
            mail.login(self.current_account["email"], self.current_account["password"])
            mail.select('inbox')

            result, data = mail.search(None, 'ALL')
            email_ids = data[0].split()

            for e_id in email_ids:
                result, msg_data = mail.fetch(e_id, '(RFC822)')
                raw_email = msg_data[0][1]
                msg = email.message_from_bytes(raw_email)
                self.emails.append((e_id, msg))

                item_text = f"Von: {msg.get('From')} - Betreff: {msg.get('Subject')}"
                item = QListWidgetItem(item_text)
                if e_id in self.marked_emails:
                    color = self.marked_emails[e_id]
                    item.setBackground(color)
                self.email_list.addItem(item)

        except Exception as e:
            QMessageBox.critical(self, "Fehler", f"E-Mail-Abruf fehlgeschlagen: {str(e)}")

    def display_email(self):
        selected_index = self.email_list.currentRow()
        if selected_index == -1:
            return

        _, selected_email = self.emails[selected_index]
        if selected_email.is_multipart():
            content = ''
            for part in selected_email.walk():
                if part.get_content_type() == "text/plain":
                    content += part.get_payload(decode=True).decode()
        else:
            content = selected_email.get_payload(decode=True).decode()

        self.email_content.setText(content)

    def send_email(self):
        if not self.current_account:
            QMessageBox.warning(self, "Warnung", "Kein E-Mail-Konto ausgewählt.")
            return

        selected_index = self.email_list.currentRow()
        if selected_index == -1:
            QMessageBox.warning(self, "Warnung", "Keine E-Mail ausgewählt.")
            return

        _, selected_email = self.emails[selected_index]
        to_address = email.utils.parseaddr(selected_email.get('From'))[1]
        subject = "Re: " + selected_email.get('Subject')
        body = self.response_text.toPlainText()

        msg = MIMEText(body)
        msg['Subject'] = subject
        msg['From'] = self.current_account["email"]
        msg['To'] = to_address

        try:
            server = smtplib.SMTP_SSL(self.current_account["smtp_server"])
            server.login(self.current_account["email"], self.current_account["password"])
            server.sendmail(self.current_account["email"], to_address, msg.as_string())
            server.quit()

            QMessageBox.information(self, "Erfolg", "E-Mail erfolgreich gesendet!")
        except Exception as e:
            QMessageBox.critical(self, "Fehler", f"E-Mail-Versand fehlgeschlagen: {str(e)}")

    def show_context_menu(self, pos):
        menu = QMenu()

        mark_action = QAction("Markieren", self)
        mark_action.triggered.connect(self.mark_email)
        menu.addAction(mark_action)

        unmark_action = QAction("Markierung aufheben", self)
        unmark_action.triggered.connect(self.unmark_email)
        menu.addAction(unmark_action)

        menu.exec_(self.email_list.viewport().mapToGlobal(pos))

    def mark_email(self):
        selected_index = self.email_list.currentRow()
        if selected_index == -1:
            return

        dialog = ColorSelectionDialog(self.color_legend, parent=self)
        if dialog.exec_() == QDialog.Accepted:
            color = dialog.selected_color()
            if color:
                email_id, _ = self.emails[selected_index]
                self.marked_emails[email_id] = color
                self.save_marked_emails()  # Speichern der Markierungen
                self.load_emails()

    def unmark_email(self):
        selected_index = self.email_list.currentRow()
        if selected_index == -1:
            return

        email_id, _ = self.emails[selected_index]
        if email_id in self.marked_emails:
            del self.marked_emails[email_id]
            self.save_marked_emails()  # Speichern der Markierungen
        self.load_emails()


# Anwendung starten
app = QApplication(sys.argv)
window = EmailClient()
window.show()
sys.exit(app.exec_())
