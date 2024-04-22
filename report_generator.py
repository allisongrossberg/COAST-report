import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime
import math

DATA_FILE = "/Users/allisongrossberg/Library/Mobile Documents/com~apple~CloudDocs/Graduate School/COAST_Study/symptom_report/COAST_Data_9_25_23.xlsx"
TEMPLATE_FILE = "coast_report_template.docx"

#List of Study IDs to generate reports for
study_ids = []

covid_symptoms = [
    "nasal_congestion",
    "sore_throat",
    "runny_nose",
    "ear_pain",
    "cough",
    "sputum_production",
    "difficulty_breathing_sob",
    "hoarse_voice",
    "chest_pain_tightness",
    "chills",
    "swollen_lymph_nodes",
    "skipping_meals_appetite_loss",
    "insomnia_sleep_problems",
    "sensitivity_heat_cold",
    "sweats",
    "white_red_purple_swollen_fingers_toes",
    "fever_feverish",
    "fatigue",
    "mood_swings_irritability",
    "weight_loss",
    "drowsiness",
    "tachycardia_arrhythmia_palpitations",
    "eye_soreness_discomfort",
    "reduced_blurred_vision",
    "photophobia_phonophobia",
    "brain_fog",
    "confusion",
    "memory_problems",
    "difficulty_concentrating",
    "delirium",
    "difficulty_finding_words",
    "paresthesia",
    "headache",
    "los",
    "lot",
    "dizziness_lightheadedness",
    "difficulty_balancing",
    "tremors",
    "stroke",
    "seizures",
    "hypoacusis",
    "numbness_hands_feet",
    "hypoethesia",
    "abdominal_pain_stomachache",
    "diarrhea",
    "nausea_vomiting",
    "muscle_weakness",
    "muscle_pains_aches",
    "bone_joint_pain",
    "neck_back_pain"
]

tbi_symptoms = [
    "headache",
    "nausea",
    "vomiting",
    "balance_problems",
    "dizziness",
    "lightheadedness",
    "fatigue",
    "trouble_falling_asleep",
    "sleeping_more",
    "sleeping_less",
    "drowsiness",
    "light_sensitivity",
    "noise_sensitivity",
    "irritability",
    "feeling_frustrated_impatient",
    "taking_longer_to_think",
    "restlessness",
    "sadness",
    "nervousness_anxiousness",
    "feeling_more_emotional",
    "numbness_tingling",
    "feeling_slowed_down",
    "in_a_fog",
    "difficulty_concentrating",
    "difficulty_remembering",
    "blurred_vision",
    "double_vision",
    "pain"
]


def format_symptom(symptom):
    return symptom.replace("_", " ").title()

def generate_report(study_id):
    data_df = pd.read_excel(DATA_FILE, index_col=False)
    participant_df = data_df[data_df["participant_id"] == study_id]

    participant_dict = participant_df.to_dict("records")[0]
    report_dict = {}

    first_name = participant_dict["qq_first_name"]
    report_dict.update({"first_name": first_name})
    last_name = participant_dict["qq_last_name"]

    print(f"Generating Report for {first_name} {last_name}")

    report_dict.update({"last_name": last_name})
    start_date = str(participant_dict["qq_start_date"]).split(" ")[0]
    report_dict.update({"start_date": start_date})

    #geographic location
    state = participant_dict["qq_state_of_residence"] if isinstance(participant_dict["qq_state_of_residence"], str) else "State of Residence Not Reported"
    report_dict.update({"state_of_residence": state})
    #employment status
    employment_dict = {
        1: "Employed for wages (part- time or full-time)",
        2: "Self-employed",
        3: "Out of work for 1 year or more",
        4: "Out of work for less than 1 year",
        5: "A homemaker",
        6: "A student",
        7: "Retired",
        8: "Unable to work (disabled)"
    }
    employment_list = []
    for i in range(1, 9, 1):
        status = participant_dict[f"qq_employment_status___{i}"]
        if status == "Checked":
            employment_list.append(employment_dict[i])
    employment_string = ", ".join(employment_list)
    if not employment_string:
        employment_string = "Employment Status Not Reported"
    report_dict.update({"employment_status": employment_string})
    #marital status
    marital_status = participant_dict["qq_marital_status"] if isinstance(participant_dict["qq_marital_status"], str) else "Marital Status Not Reported"
    report_dict.update({"marital_status": marital_status})

    # Current Diseases/Conditions
    diseases_conditions_dict = {
        "qq_anemia": "Anemia",
        "qq_asthma": "Asthma",
        "qq_copd": "COPD",
        "qq_congenital_heart_disease": "Congenital Heart Disease",
        "qq_coronary_artery_disease_history_heart_disease": "Coronary Artery Disease/Heart Disease",
        "qq_congestive_heart_failure": "Congestive Heart Failure",
        "qq_hypertension_high_bp": "Hypertension/High Blood Pressure",
        "qq_hyperlipidemia_hypercholesterolemia_high_cholesterol": "Hyperlipidemia/Hypercholesterolemia/High Cholesterol",
        "qq_liver_disease": "Liver Disease",
        "qq_type_i_diabetes": "Type 1 Diabetes",
        "qq_type_ii_diabetes": "Type 2 Diabetes",
        "qq_obesity": "Obesity",
        "qq_tick_borne_illness": ["Tick Borne Illness", "qq_tick_borne_disease_type"],
        "qq_rheumatoid_arthritis": "Rheumatoid Arthritis",
        "qq_osteoarthritis_joint_disease": "Osteoarthritis/Joint Disease",
        "qq_cystic_fibrosis": "Cystic Fibrosis",
        "qq_blood_clots": "Blood Clots",
        "qq_chronic_kidney_disease": "Chronic Kidney Disease",
        "qq_depressive_disorder": ["Depressive Disorder", "qq_depressive_disorder_type"],
        "qq_anxiety_disorder": ["Anxiety Disorder", "qq_anxiety_disorder_type"],
        "qq_adhd": "ADHD",
        "qq_bipolar_disorder": "Bipolar Disorder",
        "qq_ocd": "Obsessive Compulsive Disorder",
        "qq_ptsd": "Post-Traumatic Stress Disorder",
        "qq_schizophrenia": "Schizophrenia",
        "qq_hepatitis": "Hepatitis",
        "qq_aids_hiv": "HIV/AIDS",
        "qq_meningitis": "Meningitis",
        "qq_prion_disease": "Prion Disease",
        "qq_alzheimers_disease": "Alzheimer's Disease",
        "qq_headaches": "Headaches",
        "qq_cancer": ["Cancer", "qq_cancer_type"],
        "qq_neurological_disorder_disease_dementia": ["Neurological Disorder/Dementia", "qq_neurological_condition_type"],
        "qq_other_current_disease": ["Other", "qq_other_current_disease_type"],
    }
    condition_list = []
    for key, value in diseases_conditions_dict.items():
        if isinstance(participant_dict[key], str) and "Yes" in participant_dict[key]:
            if isinstance(value, str):
                condition_list.append(value)
            else:
                if isinstance(participant_dict[value[1]], str):
                    condition = value[0] + " (" + participant_dict[value[1]] + ")"
                else:
                    condition = value[0] + " (Type Not Specified)"
                condition_list.append(condition)
    conditions_string = ", ".join(condition_list)
    report_dict.update({"current_diseases_conditions": conditions_string})
    # Family History
    family_history_dict = {
        1: "Alzheimer's Disease",
        2: "Parkinson's Disease",
        3: "Amyotrophic Lateral Sclerosis (ALS)",
        4: "Multiple Sclerosis (MS)",
        5: "qq_family_history_disease_other",
        6: "Huntington's disease",
        7: "None of the above",
        8: "I don’t know",
    }
    family_history_list = []
    for key, value in family_history_dict.items():
        status = participant_dict[f"qq_family_history_disease___{key}"]
        if status == "Checked":
            if "_" in value:
                family_history_list.append(participant_dict[value])
            else:
                family_history_list.append(family_history_dict[key])
    family_history_string = ", ".join(family_history_list)
    report_dict.update({"family_history": family_history_string})
    # Immune Related Conditions
    immune_conditions_dict = {
    "qq_lupus": "Lupus",
    "qq_multiple_sclerosis": "Multiple Sclerosis",
    "qq_cytopenia": "Cytopenia",
    "qq_colitis_ibd": "Colitis/IBD",
    "qq_periodic_frequent_fevers": "Periodic Frequent Fevers",
    "qq_immune_deficiency": "Immune Deficiency",
    "qq_warts_skin_infections": "Warts/Skin Infections",
    "qq_allergies_hay_fever": "Allergies/Hay fever",
    "qq_food_allergies": "Food Allergies",
    "qq_cold_sores": "Cold Sores",
    "qq_shingles": "Shingles",
    "qq_eczema": "Eczema",
    "qq_hives": "Hives",
    "qq_frequent_illness": "Frequent Illness",
    "qq_thyroid_condition": ["Thyroid Condition", "qq_thyroid_condition_type"],
    "qq_other_inflammatory_condition": ["Other Inflammatory Condition", "qq_other_inflammatory_condition_type"],
    "qq_other_autoimmune_condition": ["Other Autoimmune Condition", "qq_other_autoimmune_conditon_type"],
    "qq_other_immune_related_condition": ["Other Immune Related Condition", "qq_other_immune_related_condition_type"],
    }
    immune_condition_list = []
    for key, value in immune_conditions_dict.items():
        if isinstance(participant_dict[key], str) and "Yes" in participant_dict[key]:
            if isinstance(value, str):
                immune_condition_list.append(value)
            else:
                if isinstance(participant_dict[value[1]], str):
                    condition = value[0] + " (" + participant_dict[value[1]] + ")"
                else:
                    condition = value[0] + " (Type Not Specified)"
                immune_condition_list.append(condition)
    immune_conditions_string = ", ".join(immune_condition_list)
    report_dict.update({"immune_related_conditions": immune_conditions_string})
    # Medications/Supplements
    medication_supplement_dict = {
        1: "Aspirin, with or without a prescription",
        2: "Non-steroidal anti-inflammatory agents (NSAIDS) with or without a prescription: (eg. ibuprofen (Motrin, Advil), naproxen (Naprosyn, Aleve, Anaprox, Naprelan), diclofenac (Cambia, Cataflam, Voltaren, Zipsor), indomethacin (Indocin), diflunisal, etodolac, ketoprofen, ketorolac, nambumetone, oxaprozin (Daypro), piroxicam (Feldene), salsalate (Disalate), sulidnac, tolmetin, celecoxib (Celebrex)",
        3: "Acetaminophen (Tylenol and others)",
        4: "Oral corticosteroids (eg. Prednisone)",
        5: "Inhaled corticosteroids (eg. fluticasone (Flovent), beclomethasone (QVar), etc )",
        6: "Inhaled bronchodialators (eg. albuterol)",
        7: "Other Asthma Medications",
        8: "Nerve pain medication (eg. gabapetin)",
        9: "Diabetes medication",
        10: "Anti-TNF medications (infliximab, adalimumab, certolizumab, golimumab, etanercept, others)",
        11: "IL-6 pathway inhibitors (sarilumab,tocilizumab, siltuximab, others)",
        12: "Conventional disease-modifying anti-rheumatic drugs (DMARDs) (eg. cyclosporin, cyclophosphamide, hydroxychloroquine, leflunomide, methotrexate, mycophenolate, sulfasalazine)",
        13: "JAK Inhibitors (Baricitinib, ruxolitinib, fedratinib, tofacitinib)",
        14: "Blood thinning medication (eg. warfarin (Coumadin), heparin, enoxaparin (Lovenox), apixaban (Eliquis), rivaroxaban (Xarelto), etc)",
        15: "Platelet inhibitors (eg. clopidogrel (Plavix), prasugrel (Effient), ticagrelor (Brilinta), etc.)",
        16: "Blood pressure medication: ACE inhibitors (eg. benazepril, captopril, enalapril, fosinopril, lisinopril, etc.)",
        17: "Blood pressure medication: Angiotensin Receptor Blockers (eg. losartan, valsartan, irbesartan, candesartan, telmisartan, Olmesartan, etc)",
        18: "Blood pressure medication: beta-blockers (eg. metoprolol, atenolol, carvedilol, etc.)",
        19: "Blood pressure medication: others",
        20: "Cholesterol medication: Statins (eg. atorvastatin, rosuvastatin, simvastatin, pravastatin, lovastatin, fluvastatin, pitavastatin)",
        21: "Cholesterol medication: others (ezetimibe, fenofibrate, etc)",
        22: "Thyroid medication (eg. levothryroxine, Synthroid)",
        23: "qq_current_meds_other_type",
        24: "None of the above",
    }
    medication_supplement_list = []
    for key, value in medication_supplement_dict.items():
        status = participant_dict[f"qq_current_meds___{key}"]
        if status == "Checked":
            if "_" in value:
                medication_supplement_list.append(participant_dict[value])
            else:
                medication_supplement_list.append(medication_supplement_dict[key])
    medication_supplement_string = ", ".join(medication_supplement_list)
    report_dict.update({"current_medications_supplements": medication_supplement_string})

    if (
        participant_dict["qq_covid_number"] and 
        #not math.isnan(participant_dict["qq_covid_number"]) and
        participant_dict["qq_covid_number"] != "I have never had COVID-19"
    ):
        covid_number = int(participant_dict["qq_covid_number"])
    else:
        covid_number = 0
    report_dict.update({"covid_number": covid_number})

    covid_rows = []
    if covid_number:
        for i in range(covid_number):
            row_dict = {}
            incidence_num = i + 1
            if participant_dict[f"qq_covid_{incidence_num}_symptom_status"] != "Yes":
                row_dict.update({"label": str(incidence_num), "date": "No Symptom Onset", "symptoms": "No Symptoms Reported"})
                covid_rows.append(row_dict)
            else:
                month = str(int(participant_dict[f"qq_covid_{incidence_num}_symptom_onset_month"]))
                day = str(int(participant_dict[f"qq_covid_{incidence_num}_symptom_onset_day"])) 
                year = str(int(participant_dict[f"qq_covid_{incidence_num}_symptom_onset_year"]) + 1899)
                date = f"{month}/{day}/{year}"
                symptom_list = []
                for symptom in covid_symptoms:
                    if participant_dict[f"qq_covid_{incidence_num}_duration_{symptom}"] == "I am still experiencing this symptom":
                        symptom_list.append(format_symptom(symptom))
                symptom_str = ", ".join(symptom_list)
                if not symptom_str:
                    symptom_str = "N/A - No Ongoing Symptoms"
                row_dict.update({"label": str(incidence_num), "date": date, "symptoms": symptom_str})
                covid_rows.append(row_dict)
        report_dict.update({"covid_rows": covid_rows})
    else:
        report_dict.update({"covid_rows": [{"label": "", "date": "", "symptoms": "No Reported History of COVID-19"}]})

    if participant_dict["qq_tbi_num"] and not math.isnan(participant_dict["qq_tbi_num"]):
        tbi_number = int(participant_dict["qq_tbi_num"])
    else:
        tbi_number = 0
    report_dict.update({"tbi_number": tbi_number})

    tbi_rows = []
    if tbi_number:
        for i in range(tbi_number):
            row_dict = {}
            incidence_num = i + 1
            month = str(participant_dict[f"qq_tbi_{incidence_num}_month"])
            year = str(int(participant_dict[f"qq_tbi_{incidence_num}_year"]))
            date = f"{month} {year}"
            symptom_list = []
            for symptom in tbi_symptoms:
                if participant_dict[f"qq_tbi_{incidence_num}_duration_{symptom}"] == "I am still experiencing this symptom":
                    symptom_list.append(format_symptom(symptom))
            symptom_str = ", ".join(symptom_list)
            if not symptom_str:
                symptom_str = "N/A - No Ongoing Symptoms"
            row_dict.update({"label": str(incidence_num), "date": date, "symptoms": symptom_str})
            tbi_rows.append(row_dict)
        report_dict.update({"tbi_rows": tbi_rows})
    else:
        report_dict.update({"tbi_rows": [{"label": "", "date": "", "symptoms": "No Reported History of Brain Injury"}]})

    vaccine_dose_cols = [
        "qq_covid_vaccination_doses___1",
        "qq_covid_vaccination_doses___2",
        "qq_covid_vaccination_doses___3",
        "qq_covid_vaccination_doses___4",
    ]

    label_dict = {
        "1": "First Dose",
        "2": "Second Dose",
        "3": "Third Dose",
        "4": "Booster"
    }

    vaccine_incidence_list = []
    for col in vaccine_dose_cols:
        if participant_dict[col] == "Checked":
            vaccine_incidence_list.append(col[-1])

    vaccine_number = len(vaccine_incidence_list)
    report_dict.update({"vaccine_number": vaccine_number})

    vaccine_rows = []
    if vaccine_incidence_list:
        for num in vaccine_incidence_list:
            row_dict = {}
            if num == "4":
                num =  "booster"
            if math.isnan(participant_dict[f"qq_covid_vaccination_dose_{num}_date_month"]):
                month = ""
            else:
                month = str(int(participant_dict[f"qq_covid_vaccination_dose_{num}_date_month"]))
            if math.isnan(participant_dict[f"qq_covid_vaccination_dose_{num}_date_day"]):
                day = ""
            else:
                day = str(int(participant_dict[f"qq_covid_vaccination_dose_{num}_date_day"]))
            if math.isnan(participant_dict[f"qq_covid_vaccination_dose_{num}_date_year"]):
                year = ""
            else:
                year = str(int(participant_dict[f"qq_covid_vaccination_dose_{num}_date_year"]) + 1899)
            date = f"{month}/{day}/{year}"
            if num == "4":
                vaccine_type_col = "qq_covid_vaccine_type_booster"
            else:
                vaccine_type_col = f"qq_covid_vaccine_type_dose_{num}"
            vaccine_type = participant_dict[vaccine_type_col]
            if not participant_dict[vaccine_type_col]:
                vaccine_type = participant_dict[f"qq_covid_vaccine_type_dose_{num}_other"]
            else:
                vaccine_type = participant_dict[vaccine_type_col]
            row_dict.update({"label": label_dict[num], "date": date, "type": vaccine_type})
            vaccine_rows.append(row_dict)
        report_dict.update({"vaccine_rows": vaccine_rows})
    else:
        report_dict.update({"vaccine_rows": [{"label": "", "date": "No Vaccination History", "type": ""}]})


    doc = DocxTemplate(TEMPLATE_FILE)
    doc.render(report_dict)
    today = datetime.today().strftime('%Y-%m-%d')
    report_string = f"reports/{last_name}_{first_name}_y1_qq_report_{today}.docx"
    doc.save(report_string)
    print(f"Saved {report_string}")

for study_id in study_ids:
    generate_report(study_id=study_id)