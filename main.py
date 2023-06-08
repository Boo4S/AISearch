import tkinter as tk
from tkinter import filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from tkinter import ttk
import requests
from bs4 import BeautifulSoup
import time
from requests.exceptions import SSLError, ConnectTimeout

class App:
    def __init__(self, root):
        self.root = root
        self.root.geometry("300x220")

        # Cadre pour le menu
        self.menu_frame = tk.Frame(root, width=150, bg="grey", height=50, relief='sunken')
        self.menu_frame.grid(row=0, column=0, sticky='ns')

        # Boutons du menu
        self.simple_search_button = tk.Button(self.menu_frame, text="Recherche Simple", command=self.show_simple_search)
        self.simple_search_button.pack(fill='both')

        self.identity_search_button = tk.Button(self.menu_frame, text="Recherche Identité", command=self.show_identity_search)
        self.identity_search_button.pack(fill='both')

        # Cadre pour le contenu
        self.content_frame = tk.Frame(root)
        self.content_frame.grid(row=0, column=1, sticky='nsew')

        # Sous-interfaces pour chaque type de recherche
        self.simple_search_interface = self.create_simple_search_interface()
        self.identity_search_interface = self.create_identity_search_interface()

        last_row_index = 6  # Remplacez cette valeur par l'index de la dernière ligne souhaitée.
        self.progress = ttk.Progressbar(self.simple_search_interface, orient='horizontal', length=100, mode='determinate')
        self.progress.grid(row=last_row_index, column=0)  # Utilisez last_row_index pour positionner la barre de progression.

        # Ajustement automatique de la taille des colonnes et des lignes
        root.grid_columnconfigure(1, weight=1)
        root.grid_rowconfigure(0, weight=1)

        self.df = None
        self.filename = None
        self.current_row = 0
        self.driver = webdriver.Chrome(service=Service(r'C:\Users\maxime.cedelle\Desktop\AISearch-2\chromedriver'))

    def create_simple_search_interface(self):
        frame = tk.Frame(self.content_frame)

        self.upload_button = tk.Button(frame, text="Upload Excel", command=self.upload_file)
        self.upload_button.grid(row=0, column=0)

        self.start_button = tk.Button(frame, text="Commencer la recherche", command=self.start_search, state=tk.DISABLED)
        self.start_button.grid(row=1, column=0)

        self.update_button = tk.Button(frame, text="Mise à jour Excel", command=self.update_excel)
        self.update_button.grid(row=2, column=0)

        return frame

    def create_identity_search_interface(self):
        frame = tk.Frame(self.content_frame)

        # Bouton pour uploader un fichier Excel
        self.upload_button_identity = tk.Button(frame, text="Upload Excel", command=self.upload_file)
        self.upload_button_identity.pack()

        # Zone de texte pour le nom
        self.name_label = tk.Label(frame, text="Nom")
        self.name_label.pack()
        self.name_entry = tk.Entry(frame)
        self.name_entry.pack()

        # Zone de texte pour le prénom
        self.surname_label = tk.Label(frame, text="Prénom")
        self.surname_label.pack()
        self.surname_entry = tk.Entry(frame)
        self.surname_entry.pack()

        # Checkbox pour afficher ou cacher la zone de texte pour l'année de naissance
        self.show_birth_year_check = tk.Checkbutton(frame, text="Inclure l'année de naissance", command=self.toggle_birth_year)
        self.show_birth_year_check.pack()

        # Zone de texte pour l'année de naissance (cachée par défaut)
        self.birth_year_label = tk.Label(frame, text="Année de naissance")
        self.birth_year_entry = tk.Entry(frame)
        self.birth_year_entry.pack()
        self.birth_year_label.pack()
        self.birth_year_label.pack_forget()
        self.birth_year_entry.pack_forget()

        # Bouton pour lancer la recherche
        self.start_identity_search_button = tk.Button(frame, text="Commencer la recherche", command=self.start_identity_search)
        self.start_identity_search_button.pack()

        return frame
    
    def start_identity_search(self):
        name = self.name_entry.get()
        surname = self.surname_entry.get()

        if name and surname:
            # Effectue une recherche SerpAPI pour les données entrées
            results = self.search_person(name, surname)

            # Affiche les résultats dans une fenêtre contextuelle
            self.show_results(results)
        elif self.df is not None:
            for _, row in self.df.iterrows():
                name = row['nom']
                surname = row['prenom']

                # Effectue une recherche SerpAPI pour chaque personne
                results = self.search_person(name, surname)

                # Affiche les résultats dans une fenêtre contextuelle
                self.show_results(results)

            # Affiche une pop-up pour informer l'utilisateur que toutes les recherches sont terminées
            messagebox.showinfo("Information", "Toutes les recherches sont terminées.")
        else:
            messagebox.showinfo("Information", "Veuillez d'abord uploader un fichier Excel ou entrer des données dans les champs de texte.")

    def search_person(self, name, surname):
        social_info = {"Nombre": 0, "Liens": [], "Noms": []}
        digital_life = {"Nombre": 0, "Liens": [], "Noms": []}
        digital_life_news = {"Nombre": 0, "Liens": [], "Noms": []}  # Nouvelle catégorie pour les actualités de la vie numérique
        company_info = {"Nombre": 0, "Liens": [], "Noms": []}
        company_sites = ['societe.com', 'infogreffe.fr', 'b-reputation.com', 'verif.com']

        params = {
            "engine": "google",
            "q": f"{name} {surname}",
            "api_key": "9b0d4c0366546a7bd81c14d13ae3f304ea744bff2faa67fab9eed518194b7f40",
            "hl": "fr",
            "gl": "fr",
            "google_domain": "google.com",
            "location": "France"
        }

        for i in range(2):  # limitez à 2 pages
            params["start"] = i*10

            try:
                response = requests.get('https://serpapi.com/search', params)
                data = response.json()
            except Exception as e:
                print(f"Erreur lors de la récupération des résultats de recherche : {e}")
                continue

            for result in data.get('organic_results', []):
                url = result['link']
                title = result.get('title', '').lower()

                if name.lower() in title and surname.lower() in title:
                    if 'linkedin.com' in url or 'facebook.com' in url or 'twitter.com' in url or 'instagram.com' in url or 'pinterest.com' in url or 'tiktok.com' in url:
                        social_info["Nombre"] += 1
                        social_info["Liens"].append(url)
                        social_info["Noms"].append(name + " " + surname)
                    elif any(company_site in url for company_site in company_sites):
                        company_info["Nombre"] += 1
                        company_info["Liens"].append(url)
                        company_info["Noms"].append(name + " " + surname)
                    else:
                        digital_life["Nombre"] += 1
                        digital_life["Liens"].append(url)
                        digital_life["Noms"].append(name + " " + surname)
                        
        params["tbm"] = "nws"
        params["start"] = 0 

        try:
            response = requests.get('https://serpapi.com/search', params)
            data = response.json()
        except Exception as e:
            print(f"Erreur lors de la récupération des résultats de recherche d'actualités : {e}")
            return

        for result in data.get('organic_results', []):
            url = result['link']
            title = result.get('title', '').lower()
            if f"{name.lower()} {surname.lower()}" in title:
                digital_life_news["Nombre"] += 1  # Mettez à jour la catégorie 'Vie numerique actualites'
                digital_life_news["Liens"].append(url)
                digital_life_news["Noms"].append(name + " " + surname)

        results = {
            "Reseaux sociaux": social_info,
            "Vie numerique": digital_life,
            "Vie numerique actualites": digital_life_news,  # Ajoutez cette nouvelle catégorie aux résultats
            "Entreprise": company_info
        }

        return results
    

    def show_results(self, results):
        # Créer une nouvelle fenêtre pour afficher les résultats de la recherche
        results_window = tk.Toplevel(self.root)
        results_window.title("Résultats de la recherche")

        # Créer un widget texte pour afficher les nombres de résultats
        results_text = tk.Text(results_window)
        results_text.pack()

        # Insérer les nombres de résultats dans le widget texte
        for key, value in results.items():
            results_text.insert(tk.END, f"{key}: {value['Nombre']}\n")
            detail_button = tk.Button(results_window, text=f"Voir détails de {key}", 
                                    command=lambda value=value, key=key: self.show_details(value, key))
            detail_button.pack()

        results_window.geometry("300x200")  # Ajuster la taille de la fenêtre

    def show_details(self, value, category):
        # Créer une nouvelle fenêtre pour afficher les détails
        details_window = tk.Toplevel(self.root)
        details_window.title(f"Détails de {category}")

        if 'Liens' in value:
            links_label = tk.Label(details_window, text=f"Liens:")
            links_label.pack()
            links_text = tk.Text(details_window)
            links_text.pack()
            for link in value['Liens']:
                links_text.insert(tk.END, f"{link}\n")

        if 'Noms' in value:
            names_label = tk.Label(details_window, text=f"Noms:")
            names_label.pack()
            names_text = tk.Text(details_window)
            names_text.pack()
            for name in value['Noms']:
                names_text.insert(tk.END, f"{name}\n")

        width = 600
        height = 100 + len(value.get('Liens', [])) * 20 + len(value.get('Noms', [])) * 20
        height = min(height, 800) 

        details_window.geometry(f"{width}x{height}")  # Définir la taille de la fenêtre

    def show_simple_search(self):
        self.hide_all()
        self.simple_search_interface.pack()
    
    def show_identity_search(self):
        self.hide_all()
        self.identity_search_interface.pack()

    def hide_all(self):
        self.simple_search_interface.pack_forget()
        self.identity_search_interface.pack_forget()

    def toggle_birth_year(self):
        if self.birth_year_label.winfo_ismapped():
            self.birth_year_label.pack_forget()
            self.birth_year_entry.pack_forget()
        else:
            self.birth_year_label.pack()
            self.birth_year_entry.pack()


    def upload_file(self):
        self.filename = filedialog.askopenfilename(initialdir = "/", title = "Sélectionner un fichier", filetypes = (("Excel files", "*.xlsx"), ("all files", "*.*")))
        if self.filename:
            self.df = pd.read_excel(self.filename)
            self.current_row = 0
            self.start_button['state'] = tk.NORMAL

    def start_search(self):
        if self.df is not None:
            self.progress['maximum'] = len(self.df)  # Configurer le maximum de la barre de progression
            while self.current_row < len(self.df):
                self.driver.get("https://dirigeant.societe.com/pages/recherchedir.html")
                WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, "entrepdirig")))
                self.driver.find_element(By.ID, "entrepdirig").send_keys(self.df.iloc[self.current_row]["nom"])  # 'nom'
                WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, "entreppre")))
                self.driver.find_element(By.ID, "entreppre").send_keys(self.df.iloc[self.current_row]["prenom"])  # 'prenom'

                # Insérer l'année de naissance
                WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, "entrepann")))  # "entrepann" est l'ID de l'élément de saisie de l'année de naissance
                self.driver.find_element(By.ID, "entrepann").send_keys(self.df.iloc[self.current_row]["date_naissance"])  # 'date_naissance'

                self.driver.find_element(By.XPATH, "//a[contains(text(), 'Rechercher les dirigeants')]").click()

                # Attendre que les résultats soient chargés
                try:
                    WebDriverWait(self.driver, 1).until(EC.presence_of_element_located((By.CLASS_NAME, "bloc-print")))
                except TimeoutException:
                    print("Temps d'attente dépassé en attendant le chargement des résultats. Passage à la recherche suivante.")

                try:
                    num_results_element = self.driver.find_element(By.CSS_SELECTOR, ".nombre.numdisplay")
                    num_results = int(num_results_element.text)
                except NoSuchElementException:
                    num_results = 0

                # Mettre à jour le DataFrame
                self.df.at[self.current_row, "nombre de sociétés"] = num_results  # 'nombre de sociétés'

                # Mettre à jour la barre de progression
                self.progress['value'] = self.current_row
                self.progress.update()

                # Passer à la prochaine recherche
                self.current_row += 1

            # Sauvegarder les résultats dans le fichier Excel une fois toutes les recherches terminées
            self.update_excel()

            # Reset de la barre de progression après la recherche
            self.progress['value'] = 0
            self.progress.update()

            # Afficher une pop-up pour informer l'utilisateur que toutes les recherches sont terminées
            messagebox.showinfo("Information", "Toutes les recherches sont terminées.")
        else:
            messagebox.showinfo("Information", "Veuillez d'abord uploader un fichier Excel.")

    def update_excel(self):
        if self.df is not None:
            self.df.to_excel("Resultats.xlsx", index=False)
            messagebox.showinfo("Information", "Fichier Excel mis à jour.")

root = tk.Tk()
app = App(root)
root.mainloop()
