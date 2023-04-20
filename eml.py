import tkinter as tk
from tkinter import filedialog, messagebox, ttk, Toplevel
from email.parser import BytesParser
from email.policy import default
from email.utils import parseaddr, parsedate_to_datetime
import re
import csv
import webbrowser


class FelderDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Felder auswählen")

        self.selected_options = []

        self.var_absender_email = tk.BooleanVar()
        self.var_absender_name = tk.BooleanVar()
        self.var_empfaenger_email = tk.BooleanVar()
        self.var_empfaenger_name = tk.BooleanVar()
        self.var_betreff = tk.BooleanVar()
        self.var_datum_uhrzeit = tk.BooleanVar()
        self.var_mailserver_ip = tk.BooleanVar()
        self.var_topleveldomain = tk.BooleanVar()
        self.var_domain = tk.BooleanVar()

        self.create_info_text()
        self.create_checkbuttons()
        self.create_submit_and_cancel_buttons()

    def create_info_text(self):
        info_text = tk.Label(self, text="Wähle die angezeigten Attribute aus. Wenn du nichts auswählst oder auf "
                                        "abbrechen klickst, schließt sich dieses Fenster wieder.", font=("Arial", 12))
        info_text.pack(padx=10, pady=5)

    def create_checkbuttons(self):
        checkbutton1 = tk.Checkbutton(self, text="Absender E-Mail", variable=self.var_absender_email)
        checkbutton1.pack(anchor=tk.W, padx=10, pady=5)

        checkbutton2 = tk.Checkbutton(self, text="Absender Name", variable=self.var_absender_name)
        checkbutton2.pack(anchor=tk.W, padx=10, pady=5)

        checkbutton3 = tk.Checkbutton(self, text="Empfänger E-Mail", variable=self.var_empfaenger_email)
        checkbutton3.pack(anchor=tk.W, padx=10, pady=5)

        checkbutton4 = tk.Checkbutton(self, text="Empfänger Name", variable=self.var_empfaenger_name)
        checkbutton4.pack(anchor=tk.W, padx=10, pady=5)

        checkbutton5 = tk.Checkbutton(self, text="Betreff", variable=self.var_betreff)
        checkbutton5.pack(anchor=tk.W, padx=10, pady=5)

        checkbutton6 = tk.Checkbutton(self, text="Datum und Uhrzeit", variable=self.var_datum_uhrzeit)
        checkbutton6.pack(anchor=tk.W, padx=10, pady=5)

        checkbutton7 = tk.Checkbutton(self, text="Mailserver-IP", variable=self.var_mailserver_ip)
        checkbutton7.pack(anchor=tk.W, padx=10, pady=5)

        checkbutton8 = tk.Checkbutton(self, text="Top-Level-Domain", variable=self.var_topleveldomain)
        checkbutton8.pack(anchor=tk.W, padx=10, pady=5)

        checkbutton9 = tk.Checkbutton(self, text="Domain", variable=self.var_domain)
        checkbutton9.pack(anchor=tk.W, padx=10, pady=5)

    def create_submit_and_cancel_buttons(self):
        button_frame = tk.Frame(self)
        button_frame.pack(pady=10)

        submit_button = tk.Button(button_frame, text="Weiter", command=self.submit)
        submit_button.grid(row=0, column=0, padx=5)

        cancel_button = tk.Button(button_frame, text="Abbrechen", command=self.cancel)
        cancel_button.grid(row=0, column=1, padx=5)

    def cancel(self):
        self.destroy()

    def submit(self):
        if self.var_absender_email.get():
            self.selected_options.append("Absender E-Mail")
        if self.var_absender_name.get():
            self.selected_options.append("Absender Name")
        if self.var_empfaenger_email.get():
            self.selected_options.append("Empfänger E-Mail")
        if self.var_empfaenger_name.get():
            self.selected_options.append("Empfänger Name")
        if self.var_betreff.get():
            self.selected_options.append("Betreff")
        if self.var_datum_uhrzeit.get():
            self.selected_options.append("Datum und Uhrzeit")
        if self.var_mailserver_ip.get():
            self.selected_options.append("Mailserver-IP")
        if self.var_topleveldomain.get():
            self.selected_options.append("Top-Level-Domain")
        if self.var_domain.get():
            self.selected_options.append("Domain")

        self.destroy()


class MailViewer(tk.Toplevel):
    def __init__(self, parent, file_paths, selected_options):
        super().__init__(parent)

        loading_indicator = create_loading_indicator(root, root.winfo_width() // 2 - 50, root.winfo_height() // 2 - 10)

        self.title("E-Mail-Anzeige")

        self.tree = ttk.Treeview(self, show='headings')
        self.tree.pack(expand=True, fill=tk.BOTH)

        # Spalten erstellen
        self.tree['columns'] = selected_options
        for option in selected_options:
            self.tree.column(option, anchor=tk.W, width=150)
            self.tree.heading(option, text=option)

        # E-Mails einlesen und verarbeiten
        for path in file_paths:
            email_data = self.process_email_data(path, selected_options)
            self.tree.insert('', tk.END, values=email_data)

        # Button-Frame erstellen
        button_frame = tk.Frame(self)
        button_frame.pack(padx=10, pady=10)

        # Export-Button hinzufügen
        export_button = tk.Button(button_frame, text="Exportieren als CSV", command=self.export_to_csv)
        export_button.grid(row=0, column=0, padx=5)

        # Schließen-Button hinzufügen
        close_button = tk.Button(button_frame, text="Schließen", command=self.destroy)
        close_button.grid(row=0, column=1, padx=5)

        # Hauptfenster sperren
        self.grab_set()

        loading_indicator.destroy()


    def export_to_csv(self):
        # Export-Pfad auswählen
        file_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                                 filetypes=[("CSV-Dateien", "*.csv"), ("Alle Dateien", "*.*")])
        if not file_path:
            return


        # Daten aus Treeview extrahieren
        data = []
        for item in self.tree.get_children():
            row = self.tree.item(item)["values"]
            data.append(row)

        # Daten in CSV-Datei schreiben
        with open(file_path, "w", newline="", encoding="utf-8") as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(self.tree["columns"])  # Spaltenüberschriften schreiben
            writer.writerows(data)  # Daten schreiben


    def process_email_data(self, file_path, selected_options):
        with open(file_path, "rb") as f:
            msg = BytesParser(policy=default).parse(f)

        email_data = {}
        sender_email = None

        if "Absender E-Mail" in selected_options or "Absender Name" in selected_options:
            sender_name, sender_email = parseaddr(msg["From"])
            if "Absender E-Mail" in selected_options:
                email_data["Absender E-Mail"] = sender_email
            if "Absender Name" in selected_options:
                email_data["Absender Name"] = sender_name

        if "Empfänger E-Mail" in selected_options or "Empfänger Name" in selected_options:
            recipient_name, recipient_email = parseaddr(msg["To"])
            if "Empfänger E-Mail" in selected_options:
                email_data["Empfänger E-Mail"] = recipient_email
            if "Empfänger Name" in selected_options:
                email_data["Empfänger Name"] = recipient_name

        if "Betreff" in selected_options:
            email_data["Betreff"] = msg["Subject"]

        if "Datum und Uhrzeit" in selected_options:
            date = parsedate_to_datetime(msg["Date"])
            email_data["Datum und Uhrzeit"] = date.strftime('%Y-%m-%d %H:%M:%S')

        if "Mailserver-IP" in selected_options:
            server = msg["Received"]
            server_ip = re.search(r"\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}", server)
            if server_ip:
                email_data["Mailserver-IP"] = server_ip.group()
            else:
                email_data["Mailserver-IP"] = 'Nicht gefunden'

        if "Top-Level-Domain" in selected_options:
            if sender_email is None:
                sender_name, sender_email = parseaddr(msg["From"])

            tld = self.extract_tld(sender_email)
            email_data["Top-Level-Domain"] = tld

        if "Domain" in selected_options:
            if sender_email is None:
                sender_name, sender_email = parseaddr(msg["From"])

            domain = sender_email.split('@')[-1]
            email_data["Domain"] = domain

        return [email_data.get(option, '') for option in selected_options]

    def extract_tld(self, email_address):
        domain = email_address.split('@')[-1]
        tld = domain.split('.')[-1]
        return tld


def felder_auswaehlen():
    dialog = FelderDialog(root)
    root.wait_window(dialog)
    return dialog.selected_options


def emails_auswaehlen():
    selected_options = felder_auswaehlen()

    if selected_options:
        file_paths = filedialog.askopenfilenames(filetypes=[("EML files", "*.eml")])
        if file_paths:
            viewer = MailViewer(root, file_paths, selected_options)


def open_url(url):
    webbrowser.open_new(url)


def ueber_diese_anwendung():
    about_window = tk.Toplevel()
    about_window.title("Über diese Anwendung")

    info_text = ("Mit dieser Anwendung kannst du aus eml-Dateien Daten extrahieren und "
                 "weiter verarbeiten. Mein (Sebastian) Ziel war es, aus tausenden von "
                 "E-Mails die Adressen zu filtern, um am Ende die Blacklist meines "
                 "Mailservers zu aktualisieren.")
    info_label = tk.Label(about_window, text=info_text, wraplength=400, justify=tk.LEFT)
    info_label.pack(padx=10, pady=10)

    links_frame = ttk.Frame(about_window)
    links_frame.pack(padx=10, pady=10)

    github_label = tk.Label(links_frame, text="GitHub:", font="bold")
    github_label.grid(row=0, column=0, sticky=tk.W)
    github_link = tk.Label(links_frame, text="https://github.com/basteyy", fg="blue", cursor="hand2")
    github_link.grid(row=0, column=1, sticky=tk.W)
    github_link.bind("<Button-1>", lambda e: open_url("https://github.com/basteyy"))

    website_label = tk.Label(links_frame, text="Blog:", font="bold")
    website_label.grid(row=1, column=0, sticky=tk.W)
    website_link = tk.Label(links_frame, text="https://meagainsttheweb.de/eml-tool", fg="blue", cursor="hand2")
    website_link.grid(row=1, column=1, sticky=tk.W)
    website_link.bind("<Button-1>", lambda e: open_url("https://meagainsttheweb.de/eml-tool"))

    blog_label = tk.Label(links_frame, text="Kontakt:", font="bold")
    blog_label.grid(row=2, column=0, sticky=tk.W)
    blog_link = tk.Label(links_frame, text="https://eiweleit.de", fg="blue", cursor="hand2")
    blog_link.grid(row=2, column=1, sticky=tk.W)
    blog_link.bind("<Button-1>", lambda e: open_url("https://eiweleit.de"))

    close_button = ttk.Button(about_window, text="Schließen", command=about_window.destroy)
    close_button.pack(padx=10, pady=10)



def show_loading_indicator():
    loading_window = Toplevel(root)
    loading_window.title("Lade...")
    loading_window.geometry("300x50")
    loading_window.resizable(False, False)

    progress = ttk.Progressbar(loading_window, mode='indeterminate', length=280)
    progress.pack(pady=10)
    progress.start()

    loading_window.protocol("WM_DELETE_WINDOW", disable_event)  # Verhindert das Schließen des Ladebalken-Fensters
    loading_window.grab_set()  # Deaktiviert das Hauptfenster, während der Ladebalken aktiv ist
    return loading_window, progress


def create_loading_indicator(parent, x, y):
    canvas = tk.Canvas(parent, width=100, height=20, bg="white", highlightthickness=0)
    canvas.place(x=x, y=y)

    def update_loading_indicator():
        canvas.delete("all")
        current_width = canvas.winfo_width()
        current_height = canvas.winfo_height()
        loading_width = (current_width // 4) * (update_loading_indicator.counter % 4)
        canvas.create_rectangle(0, 0, loading_width, current_height, fill="blue")
        update_loading_indicator.counter = (update_loading_indicator.counter + 1) % 4
        canvas.after(100, update_loading_indicator)

    update_loading_indicator.counter = 0
    update_loading_indicator()

    return canvas


def disable_event():
    pass


def beenden():
    root.destroy()


root = tk.Tk()
root.title("EML Tool")

# button1 = tk.Button(root, text="Felder auswählen", command=lambda: felder_auswaehlen())
# button1.pack(fill=tk.X, padx=10, pady=10)

button2 = tk.Button(root, text="E-Mails auswählen", command=emails_auswaehlen)
button2.pack(fill=tk.X, padx=10, pady=10)

button3 = tk.Button(root, text="Über diese Anwendung", command=ueber_diese_anwendung)
button3.pack(fill=tk.X, padx=10, pady=10)

button4 = tk.Button(root, text="Beenden", command=beenden)
button4.pack(fill=tk.X, padx=10, pady=10)

root.mainloop()
