import os
import re
import csv
import tkinter as tk
from tkinter import ttk, filedialog
from email.parser import BytesParser
from email.policy import default
from email.utils import parseaddr, parsedate_to_datetime

has_data = False

def extract_domain(email):
    match = re.search("@([\w.]+)", email)
    if match:
        return match.group(1)
    return None

def browse_files():
    global has_data
    file_paths = filedialog.askopenfilenames(filetypes=[("EML files", "*.eml")])
    if file_paths:
        tree.delete(*tree.get_children())
        for file_path in file_paths:
            with open(file_path, "rb") as f:
                msg = BytesParser(policy=default).parse(f)

            sender_name, sender_email = parseaddr(msg["From"])
            recipient_name, recipient_email = parseaddr(msg["To"])
            server = msg["Received"]
            date = parsedate_to_datetime(msg["Date"])
            subject = msg["Subject"]

            result_text.insert(tk.END, f"Datei: {os.path.basename(file_path)}\n")
            if sender_name_var.get():
                result_text.insert(tk.END, f"Absender Name: {sender_name}\n")
            if sender_email_var.get():
                result_text.insert(tk.END, f"Absender E-Mail: {sender_email}\n")
            if recipient_name_var.get():
                result_text.insert(tk.END, f"Empfänger Name: {recipient_name}\n")
            if recipient_email_var.get():
                result_text.insert(tk.END, f"Empfänger E-Mail: {recipient_email}\n")
            if server_var.get():
                result_text.insert(tk.END, f"Versendender Server: {server}\n")
            if date_var.get():
                result_text.insert(tk.END, f"Datum: {date}\n")
            if subject_var.get():
                result_text.insert(tk.END, f"Betreff: {subject}\n")
            result_text.insert(tk.END, "\n")
    else:
        has_data = False

def export_csv():
    global has_data
    if has_data:
        csv_file_path = filedialog.asksaveasfilename(defaultextension=".csv")
        if csv_file_path:
            headers = []
            if sender_name_var.get():
                headers.append("Absender Name")
            if sender_email_var.get():
                headers.append("Absender E-Mail")
            if recipient_name_var.get():
                headers.append("Empfänger Name")
            if recipient_email_var.get():
                headers.append("Empfänger E-Mail")
            if server_var.get():
                headers.append("Versendender Server")
            if date_var.get():
                headers.append("Datum")
            if subject_var.get():
                headers.append("Betreff")

            with open(csv_file_path, "w", newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(headers)

                for file_path in filedialog.askopenfilenames(filetypes=[("EML files", "*.eml")]):
                    with open(file_path, "rb") as f:
                        msg = BytesParser(policy=default).parse(f)

                    row = []
                    if sender_name_var.get():
                        row.append(parseaddr(msg["From"])[0])
                    if sender_email_var.get():
                        row.append(parseaddr(msg["From"])[1])
                    if recipient_name_var.get():
                        row.append(parseaddr(msg["To"])[0])
                    if recipient_email_var.get():
                        row.append(parseaddr(msg["To"])[1])
                    if server_var.get():
                        row.append(msg["Received"])
                    if date_var.get():
                        row.append(parsedate_to_datetime(msg["Date"]))
                    if subject_var.get():
                        row.append(msg["Subject"])

                    writer.writerow(row)
    else:
        tk.messagebox.showwarning("Warnung", "Es gibt keine Daten zum Exportieren.")


def extract_domains():
    global has_data
    file_paths = filedialog.askopenfilenames(filetypes=[("EML files", "*.eml")])
    if file_paths:
        for i in tree.get_children():
            tree.delete(i)

        domains = {}
        for file_path in file_paths:
            with open(file_path, "rb") as f:
                msg = BytesParser(policy=default).parse(f)

            sender_email = parseaddr(msg["From"])[1]
            domain = extract_domain(sender_email)
            if domain:
                if domain not in domains:
                    domains[domain] = 0
                domains[domain] += 1

        for domain, count in domains.items():
            tree.insert("", tk.END, values=(domain, count))

        has_data = True
    else:
        has_data = False


app = tk.Tk()
app.title("EML Viewer")

sender_name_var = tk.BooleanVar()
sender_email_var = tk.BooleanVar()
recipient_name_var = tk.BooleanVar()
recipient_email_var = tk.BooleanVar()
server_var = tk.BooleanVar()
date_var = tk.BooleanVar()
subject_var = tk.BooleanVar()

tk.Checkbutton(app, text="Absender Name", variable=sender_name_var).pack()
tk.Checkbutton(app, text="Absender E-Mail", variable=sender_email_var).pack()
tk.Checkbutton(app, text="Empfänger Name", variable=recipient_name_var).pack()
tk.Checkbutton(app, text="Empfänger E-Mail", variable=recipient_email_var).pack()
tk.Checkbutton(app, text="Versendender Server", variable=server_var).pack()
tk.Checkbutton(app, text="Datum", variable=date_var).pack()
tk.Checkbutton(app, text="Betreff", variable=subject_var).pack()

browse_button = tk.Button(text="EML-Dateien auswählen", command=browse_files)
browse_button.pack()

export_csv_button = tk.Button(text="Ergebnisse als CSV exportieren", command=export_csv)
export_csv_button.pack()

extract_domains_button = tk.Button(text="Gruppierte Domains extrahieren", command=extract_domains)
extract_domains_button.pack()

tree = ttk.Treeview(app, columns=("Domain", "Anzahl"), show="headings")
tree.heading("Domain", text="Domain")
tree.heading("Anzahl", text="Anzahl")
tree.pack()

app.mainloop()