import tkinter as tk
from tkinter import messagebox
from zk import ZK
import pandas as pd
from datetime import datetime

def extraire_donnees():
    # 1. Récupération des informations
    ip_machine = entry_ip.get().strip()
    password_str = entry_pw.get().strip()
    date_debut_str = entry_debut.get().strip()
    date_fin_str = entry_fin.get().strip()

    if not ip_machine or not date_debut_str or not date_fin_str:
        messagebox.showwarning("Champs vides", "Veuillez remplir l'IP et les dates.")
        return

    try:
        password = int(password_str) if password_str else 0
        debut = datetime.strptime(date_debut_str, "%d/%m/%Y")
        fin = datetime.strptime(date_fin_str, "%d/%m/%Y").replace(hour=23, minute=59, second=59)

        zk = ZK(ip_machine, port=4370, timeout=5, password=password)
        conn = zk.connect()
        conn.disable_device()
        attendance = conn.get_attendance()
        
        data = []
        for r in attendance:
            if debut <= r.timestamp <= fin:
                data.append({
                    "ID_Employe": r.user_id,
                    "Date_Heure": r.timestamp.strftime("%Y-%m-%d %H:%M:%S"),
                    "Etat": r.status
                })
        
        if data:
            df = pd.DataFrame(data)
            
            # --- MODIFICATION ICI : Export en CSV UTF-8 ---
            # sep=';' est souvent mieux pour l'ouverture directe dans Excel en France
            nom_export = f"Export_Pointeuse_{datetime.now().strftime('%Y%m%d_%H%M')}.csv"
            df.to_csv(nom_export, index=False, sep=';', encoding='utf-8-sig') 
            # 'utf-8-sig' permet à Excel de reconnaître l'UTF-8 automatiquement
            
            messagebox.showinfo("Succès", f"Fichier CSV créé :\n{nom_export}")
        else:
            messagebox.showwarning("Info", "Aucun passage trouvé.")
        
        conn.enable_device()
        conn.disconnect()

    except ValueError:
        messagebox.showerror("Erreur Format", "Date invalide (JJ/MM/AAAA)")
    except Exception as e:
        messagebox.showerror("Erreur", f"Problème : {e}")

# --- Interface Graphique (identique) ---
root = tk.Tk()
root.title("Extracteur ZK vers CSV")
root.geometry("350x350")

tk.Label(root, text="Adresse IP :", font=("Arial", 9, "bold")).pack(pady=5)
entry_ip = tk.Entry(root, justify='center')
entry_ip.insert(0, "192.168.100.237")
entry_ip.pack()

tk.Label(root, text="Mot de passe :", font=("Arial", 9)).pack(pady=5)
entry_pw = tk.Entry(root, justify='center', show="*")
entry_pw.insert(0, "1990")
entry_pw.pack()

tk.Label(root, text="Date début (JJ/MM/AAAA) :").pack(pady=5)
entry_debut = tk.Entry(root, justify='center')
entry_debut.insert(0, datetime.now().strftime("%d/%m/%Y"))
entry_debut.pack()

tk.Label(root, text="Date fin (JJ/MM/AAAA) :").pack(pady=5)
entry_fin = tk.Entry(root, justify='center')
entry_fin.insert(0, datetime.now().strftime("%d/%m/%Y"))
entry_fin.pack()

btn_export = tk.Button(root, text="GÉNÉRER LE FICHIER CSV", command=extraire_donnees, bg="#2980b9", fg="white", font=("Arial", 10, "bold"), height=2)
btn_export.pack(pady=20, fill='x', padx=50)

root.mainloop()