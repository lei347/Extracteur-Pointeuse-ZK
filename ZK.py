import tkinter as tk
from tkinter import messagebox
from zk import ZK
import pandas as pd
from datetime import datetime

def extraire_donnees():
    # 1. Récupération des informations saisies par l'utilisateur
    ip_machine = entry_ip.get().strip()
    password_str = entry_pw.get().strip()
    date_debut_str = entry_debut.get().strip()
    date_fin_str = entry_fin.get().strip()

    # Vérification si les champs sont remplis
    if not ip_machine or not date_debut_str or not date_fin_str:
        messagebox.showwarning("Champs vides", "Veuillez remplir l'IP et les dates.")
        return

    try:
        # 2. Préparation des paramètres
        password = int(password_str) if password_str else 0
        debut = datetime.strptime(date_debut_str, "%d/%m/%Y")
        fin = datetime.strptime(date_fin_str, "%d/%m/%Y").replace(hour=23, minute=59, second=59)

        # 3. Connexion à la pointeuse
        print(f"Tentative de connexion à {ip_machine}...")
        zk = ZK(ip_machine, port=4370, timeout=5, password=password)
        conn = zk.connect()
        
        # Bloquer temporairement la machine pour une extraction stable
        conn.disable_device()
        attendance = conn.get_attendance()
        
        # 4. Filtrage des données selon les dates
        data = []
        for r in attendance:
            if debut <= r.timestamp <= fin:
                data.append({
                    "ID_Employé": r.user_id,
                    "Date_Heure": r.timestamp.strftime("%d/%m/%Y %H:%M:%S"),
                    "Type": "Entrée/Sortie" # Simplifié pour l'utilisateur
                })
        
        # 5. Création du fichier Excel
        if data:
            df = pd.DataFrame(data)
            # Nom de fichier dynamique basé sur la date du jour
            nom_export = f"Export_Pointeuse_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            df.to_excel(nom_export, index=False)
            messagebox.showinfo("Succès", f"Fichier créé avec succès :\n{nom_export}")
        else:
            messagebox.showwarning("Info", "Aucun passage trouvé pour ces dates.")
        
        # Libérer la machine et déconnecter
        conn.enable_device()
        conn.disconnect()

    except ValueError:
        messagebox.showerror("Erreur Format", "La date doit être au format JJ/MM/AAAA (ex: 01/04/2026)")
    except Exception as e:
        messagebox.showerror("Erreur Connexion", f"Impossible de joindre la machine.\nVérifiez l'IP et le réseau.\n(Détail : {e})")

# --- Création de la Fenêtre Principale ---
root = tk.Tk()
root.title("Gestionnaire de Pointeuse ZK")
root.geometry("350x350")
root.resizable