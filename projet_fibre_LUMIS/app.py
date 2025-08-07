from flask import Flask, render_template, request, redirect, url_for, session
import json
from datetime import datetime
from dateutil.relativedelta import relativedelta  # pip install python-dateutil

app = Flask(__name__)
app.secret_key = 'lumis_fibre_optique_mdp'
USERNAME = 'lumis.user'
PASSWORD = 'lumis@2025'


# ---------------------
# FONCTIONS UTILES
# ---------------------
def charger_projets():
    try:
        with open("projets.json", "r") as f:
            return json.load(f)
    except FileNotFoundError:
        return []

def sauvegarder_projets(projets):
    with open("projets.json", "w") as f:
        json.dump(projets, f, indent=4)

# Listes fixes
CLIENTS = ["EURO FIBER", "SUDALYS", "AXIANS-NIMES", "AXIANS-GASQ", "SOGETREL", "PRIME SAS"]
TACHES = ["APS", "APD", "DOE", "CAPFT", "COMAC", "GC", "CA\\DFT", "DT", "NEXLOOP", "FIBRAGE"]
ETATS = ["À faire", "En cours", "Faite", "Bloqué"]
REALISATEURS = ["Yassine", "Omar", "Mohammed", "Rajae", "Oumaima"]

# ---------------------
# ROUTES
# ---------------------

from functools import wraps

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get("logged_in"):
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated_function


@app.route("/")
@login_required
def index():
    projets = charger_projets()
    
    # Récupérer mois choisi ou mois actuel
    mois_str = request.args.get("mois")
    if mois_str:
        try:
            mois_courant = datetime.strptime(mois_str, "%Y-%m")
        except ValueError:
            mois_courant = datetime.now().replace(day=1)
    else:
        mois_courant = datetime.now().replace(day=1)

    # Filtrer projets du mois choisi (date_reception)
    projets_mois = []
    for p in projets:
        try:
            date_rec = datetime.strptime(p["date_reception"], "%Y-%m-%d")
            if date_rec.year == mois_courant.year and date_rec.month == mois_courant.month:
                projets_mois.append(p)
        except Exception:
            continue

    # Mois précédent et suivant (format "YYYY-MM")
    mois_prec = (mois_courant - relativedelta(months=1)).strftime("%Y-%m")
    mois_suiv = (mois_courant + relativedelta(months=1)).strftime("%Y-%m")

    # Format affichage mois en français (ex: Février 2025)
    mois_courant_str = mois_courant.strftime("%B %Y").capitalize()

    return render_template(
        "index.html",
        projets=projets_mois,
        mois_courant=mois_courant_str,
        mois_prec=mois_prec,
        mois_suiv=mois_suiv
    )

@app.route("/ajouter", methods=["GET", "POST"])
@login_required
def ajouter():
    if request.method == "POST":
        projets = charger_projets()
        nouveau_projet = {
            "client": request.form["client"],
            "tache": request.form["tache"],
            "projet": request.form["projet"],
            "date_reception": request.form["date_reception"],
            "realisateur": request.form["realisateur"],
            "etat": request.form["etat"],
            "date_envoi": request.form["date_envoi"]
        }
        projets.append(nouveau_projet)
        sauvegarder_projets(projets)
        return redirect(url_for("index"))
    return render_template("add_project.html", clients=CLIENTS, taches=TACHES, etats=ETATS, realisateurs=REALISATEURS)

@app.route("/modifier/<int:index>", methods=["GET", "POST"])
@login_required
def modifier(index):
    projets = charger_projets()
    if request.method == "POST":
        projets[index] = {
            "client": request.form["client"],
            "tache": request.form["tache"],
            "projet": request.form["projet"],
            "date_reception": request.form["date_reception"],
            "realisateur": request.form["realisateur"],
            "etat": request.form["etat"],
            "date_envoi": request.form["date_envoi"]
        }
        sauvegarder_projets(projets)
        return redirect(url_for("index"))
    return render_template("modifier_project.html", projet=projets[index], index=index,
                           clients=CLIENTS, taches=TACHES, etats=ETATS, realisateurs=REALISATEURS)

@app.route("/supprimer/<int:index>")
@login_required
def supprimer(index):
    projets = charger_projets()
    projets.pop(index)
    sauvegarder_projets(projets)
    return redirect(url_for("index"))

from flask import send_file
from openpyxl import Workbook
from io import BytesIO
from datetime import datetime

@app.route("/export_excel")
@login_required
def export_excel():
    projets = charger_projets()
    
    # Récupérer le mois depuis la requête
    mois_str = request.args.get("mois")
    if mois_str:
        try:
            mois_courant = datetime.strptime(mois_str, "%Y-%m")
        except:
            return "Format de mois invalide", 400
    else:
        mois_courant = datetime.now().replace(day=1)

    # Filtrer projets du mois
    projets_mois = []
    for p in projets:
        try:
            date_rec = datetime.strptime(p["date_reception"], "%Y-%m-%d")
            if date_rec.year == mois_courant.year and date_rec.month == mois_courant.month:
                projets_mois.append(p)
        except:
            continue

    # Créer un fichier Excel en mémoire
    wb = Workbook()
    ws = wb.active
    ws.title = f"Projets {mois_courant.strftime('%B %Y')}"

    # En-têtes
    headers = ['Client', 'Tâche', 'Projet', 'Date réception', 'Réalisateur', 'État', 'Date envoi']
    ws.append(headers)

    # Données
    for projet in projets_mois:
        ws.append([
            projet.get('client', ''),
            projet.get('tache', ''),
            projet.get('projet', ''),
            projet.get('date_reception', ''),
            projet.get('realisateur', ''),
            projet.get('etat', ''),
            projet.get('date_envoi', ''),
        ])

    # Sauvegarde en mémoire
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"projets_{mois_courant.strftime('%Y_%m')}.xlsx"
    return send_file(output, as_attachment=True,
                     download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/login", methods=["GET", "POST"])
def login():
    erreur = None
    if request.method == "POST":
        username = request.form.get("username")
        password = request.form.get("password")
        if username == USERNAME and password == PASSWORD:
            session["logged_in"] = True
            return redirect(url_for("index"))
        else:
            erreur = "Nom d'utilisateur ou mot de passe incorrect."
    return render_template("login.html", erreur=erreur)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


if __name__ == "__main__":
    app.run(debug=True)
