import openpyxl
from openpyxl.styles import PatternFill
from datetime import datetime, timedelta

def create_schedule():
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = "Planning Personnel"

    # Définir les couleurs pour chaque activité (tons pastels)
    colors = {
        "Petit déjeuner": "FFE5B4",
        "Salle de sport": "E6FFE6",
        "Douche": "FFD1DC",
        "Travail": "FFB3BA",
        "Courses": "FFCCCC",
        "Repas du midi": "D7BDE2",
        "Repas du soir": "D5D8DC",
        "Sommeil": "B3E0FF",
        "Ménage": "C5E1A5",
        "Réunion familiale": "ABEBC6",
        "Réunion amis": "FAD7A0",
        "Temps libre": "FFFFFF"
    }

    # Créer l'en-tête
    days = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
    for col, day in enumerate(days, start=2):
        sheet.cell(row=1, column=col, value=day)
        sheet.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 17

    # Créer les créneaux horaires
    start_time = datetime.strptime("00:00", "%H:%M")
    for row in range(2, 50):  # 24 heures * 2 (30 min intervals)
        time = start_time + timedelta(minutes=30 * (row - 2))
        sheet.cell(row=row, column=1, value=time.strftime("%H:%M"))
    sheet.column_dimensions['A'].width = 17

    # Remplir le planning
    activities = {
        "Sommeil": {"duration": 18, "days": range(5), "start_time": "22:00"},
        "Sommeil weekend": {"duration": 24, "days": [5, 6], "start_time": "22:00"},
        "Petit déjeuner": {"duration": 1, "days": range(5), "start_time": "07:00"},
        "Petit déjeuner weekend": {"duration": 1, "days": [5, 6], "start_time": "10:00"},
        "Salle de sport": {"duration": 4, "days": range(5), "start_time": "07:30"},
        "Douche": {"duration": 1, "days": range(7), "start_time": "09:30"},
        "Courses": {"duration": 3, "days": [1, 4], "start_time": "10:00"},
        "Travail": {"duration": 16, "days": range(5), "start_time": "11:00"},
        "Repas du midi": {"duration": 2, "days": range(5), "start_time": "12:30"},
        "Repas du midi weekend": {"duration": 2, "days": [5, 6], "start_time": "13:30"},
        "Repas du soir": {"duration": 2, "days": range(7), "start_time": "19:00"},
        "Ménage": {"duration": 6, "days": [5], "start_time": "10:30"},
        "Réunion familiale": {"duration": 3, "days": [6], "start_time": "20:00"},
        "Réunion amis": {"duration": 6, "days": [4], "start_time": "20:00"}
    }

    # Initialiser toutes les cellules comme "Temps libre"
    for row in range(2, 50):
        for col in range(2, 9):
            cell = sheet.cell(row=row, column=col)
            cell.value = "Temps libre"
            cell.fill = PatternFill(start_color=colors["Temps libre"], end_color=colors["Temps libre"], fill_type="solid")

    # Fonction pour obtenir l'index de ligne à partir de l'heure
    def get_row_index(time_str):
        time = datetime.strptime(time_str, "%H:%M")
        return (time.hour * 2) + (time.minute // 30) + 2

    # Fonction pour obtenir la couleur d'une activité
    def get_activity_color(activity):
        for key in colors:
            if key in activity:
                return colors[key]
        return colors["Temps libre"]

    # Remplir le planning avec les activités
    for activity, details in activities.items():
        start_row = get_row_index(details["start_time"])
        color = get_activity_color(activity)
        base_activity = activity.split()[0] if activity != "Salle de sport" else activity
        for day in details["days"]:
            for i in range(details["duration"]):
                row = (start_row + i - 1) % 48 + 2  # Wrap around to the next day if necessary
                cell = sheet.cell(row=row, column=day + 2)
                cell.value = base_activity
                cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")

    # Calculer le temps total par activité
    total_time = {}
    for activity, details in activities.items():
        base_activity = activity.split()[0] if activity != "Salle de sport" else activity
        if base_activity not in total_time:
            total_time[base_activity] = 0
        total_time[base_activity] += details["duration"] * len(details["days"]) * 30  # en minutes

    # Calculer le temps libre total
    total_time_activities = sum(total_time.values())
    total_time["Temps libre"] = 7 * 24 * 60 - total_time_activities

    # Ajouter la liste des activités avec leurs fréquences, contraintes et temps total
    start_col = 10
    sheet.cell(row=1, column=start_col, value="Activités, fréquences, contraintes et temps total:")
    activity_details = [
        ("Petit déjeuner", "Quotidien, matin, 30 min"),
        ("Salle de sport", "5 fois par semaine, matin, 2h"),
        ("Douche", "Quotidienne, matin, 30 min (à la salle de sport ou à la maison)"),
        ("Travail", "2 x 4h quotidiennes du lundi au vendredi, commence à 11h"),
        ("Courses", "2 fois par semaine, 1h30, matin avant le travail"),
        ("Repas du midi", "Quotidien, 1h, commence à 12h30 (13h30 le weekend)"),
        ("Repas du soir", "Quotidien, 1h"),
        ("Sommeil", "9h par nuit (jusqu'à 10h le weekend)"),
        ("Ménage", "3h / semaine, samedi matin avant le repas du midi"),
        ("Réunion familiale", "1h30 / semaine, dimanche à 20h"),
        ("Réunion amis", "3h / semaine, vendredi soir"),
        ("Temps libre", "Plages de temps non affectées")
    ]
    for row, (activity, details) in enumerate(activity_details, start=2):
        sheet.cell(row=row, column=start_col).fill = PatternFill(start_color=colors[activity], end_color=colors[activity], fill_type="solid")
        sheet.cell(row=row, column=start_col+1, value=activity)
        sheet.cell(row=row, column=start_col+2, value=details)
        if activity in total_time:
            hours, minutes = divmod(total_time[activity], 60)
            sheet.cell(row=row, column=start_col+3, value=f"{hours}h{minutes:02d}")

    # Ajuster la largeur des colonnes
    for col in range(start_col, start_col+4):
        sheet.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 30

    # Sauvegarder le fichier
    wb.save("planning_personnel.xlsx")

create_schedule()