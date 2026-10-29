Attribute VB_Name = "mod_TextCatalogWBSTutorials"
Option Explicit

'------------------------------------------------------------------------------
' FR: Fournit les 37 tutoriels historiques au provider TextCatalog WBS.
' EN: Supplies the 37 historical tutorials to the WBS TextCatalog provider.
'------------------------------------------------------------------------------
Public Function TextCatalogWBS_TutorialSourceMap() As Object

    Dim comments As Object

    Set comments = CreateObject("Scripting.Dictionary")
    comments.CompareMode = vbTextCompare

    TextCatalogWBS_AddTutorial comments, VTS_COL_ID, _
        "Rôle : identifiant unique de la ligne.{NL}Saisie attendue : entier unique, sans doublon.{NL}Exemple : 1 / 17 / 29{NL}{NL}Utilité :{NL}- Sert de clé technique entre WBS, CALC, GANTT et S-Curve.{NL}- Ne doit jamais être dupliqué.{NL}- Peut rester simple même si le WBS change.", _
        "Purpose: unique row identifier.{NL}Expected input: unique integer, no duplicates.{NL}Example: 1 / 17 / 29{NL}{NL}Use:{NL}- Technical key used across WBS, CALC, GANTT and S-Curve.{NL}- Must never be duplicated.{NL}- Can stay stable even if the WBS code changes."
    TextCatalogWBS_AddTutorial comments, VTS_COL_WBS, _
        "Rôle : code hiérarchique de la tâche dans la structure du projet.{NL}Saisie attendue : format numérique hiérarchique avec points.{NL}Exemple : 1.0 / 1.3 / 1.3.2 / 1.3.2.1{NL}{NL}Règles :{NL}- Utiliser uniquement des chiffres et des points.{NL}- Pas d’espace, pas de lettres.{NL}- Une tâche parent a des enfants dont le WBS commence par son propre code.", _
        "Purpose: hierarchical code of the task in the project structure.{NL}Expected input: numeric hierarchical format using dots.{NL}Example: 1.0 / 1.3 / 1.3.2 / 1.3.2.1{NL}{NL}Rules:{NL}- Use numbers and dots only.{NL}- No spaces, no letters.{NL}- A parent task has child tasks whose WBS starts with its own code."
    TextCatalogWBS_AddTutorial comments, VTS_COL_TASK_NAME, _
        "Rôle : nom court de la tâche.{NL}Saisie attendue : intitulé clair et lisible.{NL}Exemple : Kick-off / RFQ / Assembly / FAT{NL}{NL}Conseil :{NL}- Rester court.{NL}- Utiliser un nom orienté action ou livrable.", _
        "Purpose: short task name.{NL}Expected input: clear readable label.{NL}Example: Kick-off / RFQ / Assembly / FAT{NL}{NL}Tip:{NL}- Keep it short.{NL}- Prefer action-oriented or deliverable-oriented wording."
    TextCatalogWBS_AddTutorial comments, VTS_COL_TASK_DESCRIPTION, _
        "Rôle : description détaillée de la tâche.{NL}Saisie attendue : phrase courte ou précision utile.{NL}Exemple : Project official start meeting / Vendor docs review{NL}{NL}Utilité :{NL}- Aide à comprendre la tâche sans lire tout le planning.", _
        "Purpose: detailed task description.{NL}Expected input: short sentence or useful clarification.{NL}Example: Project official start meeting / Vendor docs review{NL}{NL}Use:{NL}- Helps understand the task without reading the full schedule."
    TextCatalogWBS_AddTutorial comments, VTS_COL_DISCIPLINE, _
        "Rôle : discipline technique principale concernée.{NL}Saisie attendue : nom de discipline cohérent sur tout le fichier.{NL}Exemple : Process / Mechanical / Structure / QAQC / Logistics{NL}{NL}Conseil :{NL}- Garder un vocabulaire homogène dans tout le planning.", _
        "Purpose: main technical discipline involved.{NL}Expected input: discipline name used consistently across the file.{NL}Example: Process / Mechanical / Structure / QAQC / Logistics{NL}{NL}Tip:{NL}- Keep the wording consistent throughout the schedule."
    TextCatalogWBS_AddTutorial comments, VTS_COL_SUPPLIER, _
        "Rôle : acteur principal responsable ou concerné par la tâche.{NL}Saisie attendue : nom de société, fournisseur ou entité interne.{NL}Exemple : Internal / Vendor A / Client / Forwarder", _
        "Purpose: main party responsible for or involved in the task.{NL}Expected input: company name, supplier or internal entity.{NL}Example: Internal / Vendor A / Client / Forwarder"
    TextCatalogWBS_AddTutorial comments, VTS_COL_PROJECT, _
        "Rôle : projet ou regroupement métier facultatif associé à la tâche.{NL}Saisie attendue : nom du projet ou regroupement utile.{NL}Exemple : Projet A / Fabrication / Zone 2{NL}{NL}Utilité :{NL}- Facilite le regroupement métier des tâches.{NL}- Peut rester vide.{NL}- N'intervient pas dans le calcul du planning.", _
        "Purpose: optional project or business grouping associated with the task.{NL}Expected input: project name or useful grouping.{NL}Example: Project A / Fabrication / Area 2{NL}{NL}Use:{NL}- Helps group tasks for business reporting.{NL}- May be left blank.{NL}- Does not affect schedule calculation."
    TextCatalogWBS_AddTutorial comments, VTS_COL_TASK_TYPE, _
        "Rôle : définit le comportement de la tâche dans le moteur de planning.{NL}{NL}Valeurs autorisées :{NL}Task{NL}Milestone{NL}Level of Effort{NL}{NL}Définition :{NL}{NL}Task{NL}tâche standard avec durée{NL}peut avoir des prédécesseurs{NL}pilotée par Actual / Forecast / Baseline / Dependencies{NL}impacte le réseau normalement{NL}{NL}Milestone{NL}tâche sans durée (événement ponctuel){NL}durée nulle ou minimale selon convention{NL}peut avoir des prédécesseurs{NL}représente un jalon (début / fin / validation){NL}{NL}Level of Effort{NL}tâche dépendante d’une plage d’activités{NL}ne pilote pas le planning{NL}est pilotée par ses dépendances{NL}généralement définie par SS (début) et FF (fin){NL}durée déduite du réseau{NL}{NL}Règles :{NL}une seule valeur par tâche{NL}respecter strictement les valeurs autorisées{NL}une LOE ne doit pas être utilisée comme driver d’autres tâches{NL}une milestone ne doit pas porter de durée métier", _
        "Purpose: defines the task behavior within the planning engine.{NL}{NL}Allowed values:{NL}Task{NL}Milestone{NL}Level of Effort{NL}{NL}Definition:{NL}{NL}Task{NL}standard task with duration{NL}can have predecessors{NL}driven by Actual / Forecast / Baseline / Dependencies{NL}fully participates in the network{NL}{NL}Milestone{NL}zero-duration task (event){NL}duration is zero or minimal depending on convention{NL}can have predecessors{NL}represents a key event (start / finish / validation){NL}{NL}Level of Effort{NL}task spanning a range of activities{NL}does not drive the schedule{NL}is driven by its dependencies{NL}typically defined using SS (start) and FF (finish) links{NL}duration is derived from the network{NL}{NL}Rules:{NL}one value per task{NL}must match allowed values exactly{NL}a LOE must not be used as a driver for other tasks{NL}a milestone must not carry business duration"
    TextCatalogWBS_AddTutorial comments, VTS_COL_S, _
        "Rôle : définit si la tâche doit apparaître dans la vue Summary du Gantt.{NL}Valeurs autorisées :{NL}Y{NL}N{NL}Définition :{NL}Y{NL}la tâche est affichée dans la vue Summary{NL}peut être utilisé pour afficher une tâche standard importante{NL}permet de forcer l’affichage d’une ligne même si ce n’est pas un parent ou une milestone{NL}N{NL}la tâche est masquée dans la vue Summary{NL}permet de masquer une milestone ou une ligne non pertinente{NL}n’impacte pas le calcul planning{NL}Règles :{NL}si vide, la valeur est remplie automatiquement{NL}parents / summaries : Y par défaut{NL}milestones : Y par défaut{NL}tasks standard : N par défaut{NL}Level of Effort : N par défaut{NL}une valeur déjà renseignée n’est jamais écrasée{NL}Y ou N uniquement", _
        "Purpose: defines whether the task should appear in the Gantt Summary view.{NL}Allowed values:{NL}Y{NL}N{NL}Definition:{NL}Y{NL}task is displayed in the Summary view{NL}can be used to show an important standard task{NL}forces a row to appear even if it is not a parent or milestone{NL}N{NL}task is hidden from the Summary view{NL}can be used to hide a milestone or non-relevant row{NL}does not impact schedule calculation{NL}Rules:{NL}if blank, the value is filled automatically{NL}parents / summaries: Y by default{NL}milestones: Y by default{NL}standard tasks: N by default{NL}Level of Effort: N by default{NL}an existing value is never overwritten{NL}Y or N only"
    TextCatalogWBS_AddTutorial comments, VTS_COL_COMMENTS, _
        "Rôle : zone libre pour note de contexte.{NL}Saisie attendue : commentaire court, risque, hypothèse, précision.{NL}Exemple : Waiting vendor confirmation / Milestone imposed by client", _
        "Purpose: free text field for context notes.{NL}Expected input: short comment, risk, assumption or clarification.{NL}Example: Waiting vendor confirmation / Milestone imposed by client"
    TextCatalogWBS_AddTutorial comments, VTS_COL_PREDECESSORS_WBS, _
        "Rôle : antécédents de la tâche, saisis au format WBS avec type de lien et lag éventuel.{NL}Saisie attendue :{NL}un ou plusieurs prédécesseurs séparés par un point-virgule{NL}type par défaut = FS si rien n’est précisé{NL}Exemples :{NL}1.2.3{NL}1.2.3+4{NL}1.2.3-2{NL}1.2.3FS+4{NL}1.2.3SS-2{NL}1.2.3FF{NL}1.2.3;1.4.1SS+2;2.3FF-1{NL}Règles :{NL}pas d’espace{NL}utiliser le WBS, pas l’ID{NL}types autorisés : FS, SS, FF{NL}lag autorisé en positif ou négatif{NL}le moteur convertit ensuite cette donnée en IDs techniques + table de liens logiques", _
        "Purpose: task predecessors entered using WBS codes, with optional link type and lag.{NL}Expected input:{NL}one or more predecessors separated by semicolons{NL}default link type = FS when omitted{NL}Examples:{NL}1.2.3{NL}1.2.3+4{NL}1.2.3-2{NL}1.2.3FS+4{NL}1.2.3SS-2{NL}1.2.3FF{NL}1.2.3;1.4.1SS+2;2.3FF-1{NL}Rules:{NL}no spaces{NL}use WBS, not ID{NL}allowed link types: FS, SS, FF{NL}positive or negative lag allowed{NL}the engine then converts this input into technical IDs + logical link table"
    TextCatalogWBS_AddTutorial comments, VTS_COL_WEIGHT_PERCENT, _
        "Rôle : poids de la tâche pour les analyses de charge ou de progression pondérée.{NL}Saisie attendue : valeur de poids selon la logique projet.{NL}Exemple : 20000 € / 15 / 4.5{NL}{NL}Utilité :{NL}- Peut représenter un coût, une charge, un volume ou tout autre poids relatif.{NL}- La S-Curve travaille sur les tâches feuilles uniquement et normalise ensuite les poids.", _
        "Purpose: task weight used for workload or weighted progress analysis.{NL}Expected input: weight value according to the project logic.{NL}Example: 20000 € / 15 / 4.5{NL}{NL}Use:{NL}- Can represent cost, effort, quantity or any relative weighting.{NL}- The S-Curve works on leaf tasks only and then normalizes weights."
    TextCatalogWBS_AddTutorial comments, VTS_COL_PROGRESS_PERCENT, _
        "Rôle : avancement manuel de la tâche.{NL}Saisie attendue : pourcentage entre 0% et 100%.{NL}Exemple : 0% / 8% / 70% / 100%{NL}{NL}Règles :{NL}- À renseigner sur les tâches feuilles.{NL}- En l’absence de valeur, l’affichage Gantt peut considérer 0%.", _
        "Purpose: manual task progress.{NL}Expected input: percentage between 0% and 100%.{NL}Example: 0% / 8% / 70% / 100%{NL}{NL}Rules:{NL}- Meant for leaf tasks.{NL}- When empty, Gantt display may treat it as 0%."
    TextCatalogWBS_AddTutorial comments, VTS_COL_BASELINE_START, _
        "Rôle : date de début de référence.{NL}Saisie attendue : date baseline prévue au plan initial.{NL}Exemple : 05/02/2026{NL}{NL}Utilité :{NL}- Sert de base de comparaison pour les écarts.{NL}- Utilisée par le moteur et par les analyses REX.", _
        "Purpose: reference start date.{NL}Expected input: baseline start date from the initial plan.{NL}Example: 05/02/2026{NL}{NL}Use:{NL}- Used as comparison basis for variances.{NL}- Used by the engine and by REX analyses."
    TextCatalogWBS_AddTutorial comments, VTS_COL_BASELINE_DURATION, _
        "Rôle : durée baseline en jours calendaires inclusifs.{NL}Saisie attendue : entier positif.{NL}Exemple : 1 / 5 / 12{NL}{NL}Règle importante :{NL}- Une durée de 1 jour signifie début = fin le même jour.", _
        "Purpose: baseline duration in inclusive calendar days.{NL}Expected input: positive integer.{NL}Example: 1 / 5 / 12{NL}{NL}Important rule:{NL}- A duration of 1 day means start = finish on the same day."
    TextCatalogWBS_AddTutorial comments, VTS_COL_BASELINE_FINISH, _
        "Rôle : date de fin baseline.{NL}Calcul / logique :{NL}- Colonne calculée automatiquement à partir de Baseline Start et Baseline Duration.{NL}- Logique inclusive : Finish = Start + Duration - 1{NL}{NL}Utilité :{NL}- Sert aux écarts de fin et aux comparaisons planning.", _
        "Purpose: baseline finish date.{NL}Calculation / logic:{NL}- Automatically calculated from Baseline Start and Baseline Duration.{NL}- Inclusive logic: Finish = Start + Duration - 1{NL}{NL}Use:{NL}- Used for finish variance and schedule comparisons."
    TextCatalogWBS_AddTutorial comments, VTS_COL_ACTUAL_START, _
        "Rôle : date de début réellement constatée.{NL}Saisie attendue : date réelle si la tâche a commencé.{NL}Exemple : 21/03/2026{NL}{NL}Utilité :{NL}- Prioritaire sur Forecast et Baseline pour le calcul moteur.", _
        "Purpose: actual observed start date.{NL}Expected input: real start date if the task has started.{NL}Example: 21/03/2026{NL}{NL}Use:{NL}- Has priority over Forecast and Baseline in the engine logic."
    TextCatalogWBS_AddTutorial comments, VTS_COL_ACTUAL_FINISH, _
        "Rôle : date de fin réellement constatée.{NL}Saisie attendue : date réelle si la tâche est terminée.{NL}Exemple : 23/03/2026{NL}{NL}Utilité :{NL}- Prioritaire pour le calcul de la fin si présente.", _
        "Purpose: actual observed finish date.{NL}Expected input: real finish date if the task is completed.{NL}Example: 23/03/2026{NL}{NL}Use:{NL}- Has priority for finish calculation when present."
    TextCatalogWBS_AddTutorial comments, VTS_COL_ACTUAL_DURATION, _
        "Rôle : durée réelle observée.{NL}Calcul / logique :{NL}- Colonne calculée automatiquement à partir de Actual Start et Actual Finish.{NL}- Logique inclusive : Duration = Finish - Start + 1{NL}{NL}Utilité :{NL}- Donne la durée réelle constatée sans saisie manuelle.", _
        "Purpose: actual observed duration.{NL}Calculation / logic:{NL}- Automatically calculated from Actual Start and Actual Finish.{NL}- Inclusive logic: Duration = Finish - Start + 1{NL}{NL}Use:{NL}- Provides the real duration without manual entry."
    TextCatalogWBS_AddTutorial comments, VTS_COL_FORECAST_START, _
        "Rôle : date de début prévisionnelle mise à jour.{NL}Saisie attendue : date forecast si la tâche n’est pas entièrement portée par l’Actual.{NL}Exemple : 04/04/2026{NL}{NL}Utilité :{NL}- Permet de simuler ou piloter une dérive planning.{NL}- Si incohérente avec les dépendances, le moteur bloque.", _
        "Purpose: updated forecast start date.{NL}Expected input: forecast date when the task is not fully driven by Actual data.{NL}Example: 04/04/2026{NL}{NL}Use:{NL}- Allows schedule drift management and simulation.{NL}- If inconsistent with dependencies, the engine blocks."
    TextCatalogWBS_AddTutorial comments, VTS_COL_FORECAST_FINISH, _
        "Rôle : date de fin prévisionnelle mise à jour.{NL}Saisie attendue : date forecast de fin.{NL}Exemple : 30/04/2026{NL}{NL}Utilité :{NL}- Permet d’imposer une fin forecast.{NL}- Si seule la date de début est donnée, le moteur conserve la durée de référence.", _
        "Purpose: updated forecast finish date.{NL}Expected input: forecast finish date.{NL}Example: 30/04/2026{NL}{NL}Use:{NL}- Allows forcing a forecast finish.{NL}- If only the start is given, the engine keeps the reference duration."
    TextCatalogWBS_AddTutorial comments, VTS_COL_CALCULATED_START, _
        "Rôle : date de début calculée par le moteur.{NL}Calcul / logique :{NL}- Priorité générale : Actual > Forecast > Baseline > Dépendances seules{NL}- Le moteur tient compte des prédécesseurs et du lag.{NL}{NL}Utilité :{NL}- Référence consolidée utilisée pour le Gantt et les analyses.", _
        "Purpose: engine-calculated start date.{NL}Calculation / logic:{NL}- General priority: Actual > Forecast > Baseline > Dependencies only{NL}- The engine also applies predecessors and lag.{NL}{NL}Use:{NL}- Consolidated reference used by Gantt and analyses."
    TextCatalogWBS_AddTutorial comments, VTS_COL_CALCULATED_FINISH, _
        "Rôle : date de fin calculée par le moteur.{NL}Calcul / logique :{NL}- Basée sur Actual Finish si présent, sinon Forecast Finish si présent, sinon durée de référence.{NL}- Toujours cohérente avec Calculated Start si le calcul réussit.{NL}{NL}Utilité :{NL}- Référence consolidée utilisée pour le Gantt et les analyses.", _
        "Purpose: engine-calculated finish date.{NL}Calculation / logic:{NL}- Based on Actual Finish if present, otherwise Forecast Finish if present, otherwise reference duration.{NL}- Always aligned with Calculated Start if the calculation succeeds.{NL}{NL}Use:{NL}- Consolidated reference used by Gantt and analyses."
    TextCatalogWBS_AddTutorial comments, VTS_COL_CALCULATED_DURATION, _
        "Rôle : durée calculée consolidée.{NL}Calcul / logique :{NL}- Colonne calculée automatiquement à partir de Calculated Start et Calculated Finish.{NL}- Logique inclusive : Duration = Finish - Start + 1{NL}{NL}Utilité :{NL}- Affiche la durée réellement retenue après calcul.", _
        "Purpose: consolidated calculated duration.{NL}Calculation / logic:{NL}- Automatically calculated from Calculated Start and Calculated Finish.{NL}- Inclusive logic: Duration = Finish - Start + 1{NL}{NL}Use:{NL}- Shows the duration finally retained after calculation."
    TextCatalogWBS_AddTutorial comments, VTS_COL_START_VARIANCE, _
        "Rôle : écart entre le début calculé et le début baseline.{NL}Calcul / logique :{NL}- Start Variance = Calculated Start - Baseline Start{NL}{NL}Lecture :{NL}- 0 = conforme baseline{NL}- > 0 = démarrage plus tardif{NL}- < 0 = démarrage plus tôt", _
        "Purpose: variance between calculated start and baseline start.{NL}Calculation / logic:{NL}- Start Variance = Calculated Start - Baseline Start{NL}{NL}Reading:{NL}- 0 = aligned with baseline{NL}- > 0 = later start{NL}- < 0 = earlier start"
    TextCatalogWBS_AddTutorial comments, VTS_COL_FINISH_VARIANCE, _
        "Rôle : écart entre la fin calculée et la fin baseline.{NL}Calcul / logique :{NL}- Finish Variance = Calculated Finish - Baseline Finish{NL}{NL}Lecture :{NL}- 0 = conforme baseline{NL}- > 0 = fin plus tardive{NL}- < 0 = fin plus tôt", _
        "Purpose: variance between calculated finish and baseline finish.{NL}Calculation / logic:{NL}- Finish Variance = Calculated Finish - Baseline Finish{NL}{NL}Reading:{NL}- 0 = aligned with baseline{NL}- > 0 = later finish{NL}- < 0 = earlier finish"
    TextCatalogWBS_AddTutorial comments, VTS_COL_DURATION_VARIANCE, _
        "Rôle : écart entre la durée calculée et la durée baseline.{NL}Calcul / logique :{NL}- Duration Variance = Calculated Duration - Baseline Duration{NL}{NL}Lecture :{NL}- 0 = durée inchangée{NL}- > 0 = durée plus longue{NL}- < 0 = durée plus courte", _
        "Purpose: variance between calculated duration and baseline duration.{NL}Calculation / logic:{NL}- Duration Variance = Calculated Duration - Baseline Duration{NL}{NL}Reading:{NL}- 0 = unchanged duration{NL}- > 0 = longer duration{NL}- < 0 = shorter duration"
    TextCatalogWBS_AddTutorial comments, VTS_COL_DRIVING_LOGIC, _
        "Rôle : source principale ayant piloté le calcul de la tâche.{NL}Valeurs typiques :{NL}- ACTUAL{NL}- FORECAST{NL}- BASELINE{NL}- DEPENDENCY{NL}- SUMMARY{NL}{NL}Utilité :{NL}- Permet de comprendre rapidement pourquoi la date calculée est celle-ci.", _
        "Purpose: main source driving the task calculation.{NL}Typical values:{NL}- ACTUAL{NL}- FORECAST{NL}- BASELINE{NL}- DEPENDENCY{NL}- SUMMARY{NL}{NL}Use:{NL}- Quickly explains why the calculated date is what it is."
    TextCatalogWBS_AddTutorial comments, VTS_COL_CRITICAL_PATH, _
        "Rôle : indicateur de chemin critique sur le réseau de planning actuel.{NL}Calcul / logique :{NL}- Basé sur le float total courant.{NL}- Une tâche est critique si son Total Float est inférieur ou égal à 0.{NL}{NL}Utilité :{NL}- Aide à identifier les tâches qui pilotent directement la date projet actuelle.", _
        "Purpose: critical path indicator on the current schedule network.{NL}Calculation / logic:{NL}- Based on current total float.{NL}- A task is critical if its Total Float is less than or equal to 0.{NL}{NL}Use:{NL}- Helps identify tasks directly driving the current project finish date."
    TextCatalogWBS_AddTutorial comments, VTS_COL_CRITICAL_PATH_REX, _
        "Rôle : indicateur de chemin critique sur l'état de référence Baseline.{NL}Calcul / logique :{NL}- Utilise le même graphe que les analytics Current.{NL}- Analyse les dates Baseline posées, sans compresser les gaps non exprimés en lag.{NL}- Une tâche est critique si son Total Float REX est inférieur ou égal à 0.{NL}{NL}Utilité :{NL}- Sert à l'analyse rétrospective ou comparative sur l'état Baseline.", _
        "Purpose: critical path indicator on the Baseline reference state.{NL}Calculation / logic:{NL}- Uses the same graph as Current analytics.{NL}- Analyzes the placed Baseline dates without compressing gaps that are not expressed as lag.{NL}- A task is critical if its Total Float REX is less than or equal to 0.{NL}{NL}Use:{NL}- Used for retrospective or comparative analysis on the Baseline state."
    TextCatalogWBS_AddTutorial comments, VTS_COL_LONGEST_PATH, _
        "Rôle : indique si la tâche appartient au plus long chemin du réseau de planning actuel.{NL}Calcul / logique :{NL}- Le moteur marque LONGEST les tâches non terminées reliées à la date de fin du réseau courant par des liens directeurs.{NL}{NL}Utilité :{NL}- Identifie la séquence active la plus longue jusqu'à la fin du projet.", _
        "Purpose: indicates whether the task belongs to the longest path in the current schedule network.{NL}Calculation / logic:{NL}- The engine marks unfinished tasks as LONGEST when driving links connect them to the current network finish.{NL}{NL}Use:{NL}- Identifies the longest active sequence leading to project completion."
    TextCatalogWBS_AddTutorial comments, VTS_COL_LONGEST_PATH_REX, _
        "Rôle : sortie calculée réservée au plus long chemin sur l'état de référence Baseline / REX.{NL}{NL}Règles :{NL}- Ne pas renseigner manuellement.{NL}- Peut rester vide lorsque le workflow REX ne produit pas cet indicateur.", _
        "Purpose: calculated output reserved for the longest path on the Baseline / REX reference state.{NL}{NL}Rules:{NL}- Do not enter a value manually.{NL}- May remain blank when the REX workflow does not produce this indicator."
    TextCatalogWBS_AddTutorial comments, VTS_COL_TOTAL_FLOAT, _
        "Rôle : marge totale sur le planning actuel.{NL}Calcul / logique :{NL}- Nombre de jours pendant lesquels la tâche peut glisser sans décaler la date de fin projet actuelle.{NL}{NL}Lecture :{NL}- 0 = critique{NL}- > 0 = marge disponible{NL}- < 0 = float négatif, planning incohérent ou contraint", _
        "Purpose: total float on the current schedule.{NL}Calculation / logic:{NL}- Number of days the task can slip without delaying the current project finish date.{NL}{NL}Reading:{NL}- 0 = critical{NL}- > 0 = available margin{NL}- < 0 = negative float, constrained or inconsistent schedule"
    TextCatalogWBS_AddTutorial comments, VTS_COL_FREE_FLOAT, _
        "Rôle : marge libre sur le planning actuel.{NL}{NL}Calcul / logique :{NL}- Nombre de jours pendant lesquels la tâche peut glisser sans impacter le début au plus tôt de la tâche suivante.{NL}{NL}Cas particulier :{NL}- Si le float est négatif, cela signifie que la tâche ne respecte pas les contraintes du réseau.{NL}- La date de fin actuelle est déjà trop tardive par rapport aux exigences des successeurs (ou la durée est insuffisante).{NL}{NL}Utilité :{NL}- Mesure la marge locale, plus fine que le Total Float.{NL}- Permet d’identifier immédiatement les incohérences ou contraintes impossibles dans le planning.", _
        "Purpose: free float on the current schedule.{NL}{NL}Calculation / logic:{NL}- Number of days the task can slip without affecting the earliest start of the next task.{NL}{NL}Special case:{NL}- A negative float means the task violates network constraints.{NL}- The current finish is already too late relative to successor requirements (or duration is insufficient).{NL}{NL}Use:{NL}- Measures local margin, more granular than Total Float.{NL}- Helps detect inconsistencies or infeasible constraints in the schedule."
    TextCatalogWBS_AddTutorial comments, VTS_COL_TOTAL_FLOAT_REX, _
        "Rôle : marge totale sur l'état de référence Baseline / REX.{NL}Calcul / logique :{NL}- Equivalent Baseline du Total Float, calculé avec le même graphe que Current sur les dates Baseline posées.{NL}{NL}Utilité :{NL}- Sert aux comparaisons et au retour d'expérience.", _
        "Purpose: total float on the Baseline / REX reference state.{NL}Calculation / logic:{NL}- Baseline equivalent of Total Float, calculated with the same graph as Current on the placed Baseline dates.{NL}{NL}Use:{NL}- Used for comparisons and lessons learned analysis."
    TextCatalogWBS_AddTutorial comments, VTS_COL_FREE_FLOAT_REX, _
        "Rôle : marge libre sur l'état de référence Baseline / REX.{NL}Calcul / logique :{NL}- Equivalent Baseline du Free Float, calculé avec le même graphe que Current sur les dates Baseline posées.{NL}{NL}Utilité :{NL}- Sert à l'analyse fine des marges sur l'état Baseline.", _
        "Purpose: free float on the Baseline / REX reference state.{NL}Calculation / logic:{NL}- Baseline equivalent of Free Float, calculated with the same graph as Current on the placed Baseline dates.{NL}{NL}Use:{NL}- Used for detailed float analysis on the Baseline state."
    TextCatalogWBS_AddTutorial comments, VTS_COL_DEADLINE_FLOAT, _
        "Rôle : marge entre la deadline de la tâche et sa date de fin calculée.{NL}Calcul / logique :{NL}- Deadline Float = Deadline - Calculated Finish.{NL}{NL}Lecture :{NL}- 0 = échéance atteinte exactement.{NL}- > 0 = marge disponible avant l'échéance.{NL}- < 0 = échéance dépassée.", _
        "Purpose: margin between the task deadline and its calculated finish date.{NL}Calculation / logic:{NL}- Deadline Float = Deadline - Calculated Finish.{NL}{NL}Reading:{NL}- 0 = deadline met exactly.{NL}- > 0 = available margin before the deadline.{NL}- < 0 = deadline exceeded."

    Set TextCatalogWBS_TutorialSourceMap = comments

End Function

'------------------------------------------------------------------------------
' FR: Ajoute une paire de textes localisés au catalogue sans exposer son stockage.
' EN: Adds one localized text pair to the catalog without exposing its storage.
'------------------------------------------------------------------------------
Private Sub TextCatalogWBS_AddTutorial( _
    ByVal comments As Object, _
    ByVal columnName As String, _
    ByVal frText As String, _
    ByVal enText As String)

    comments(columnName) = Array( _
        TextCatalogWBS_DecodeTutorialText(frText), _
        TextCatalogWBS_DecodeTutorialText(enText))

End Sub

'------------------------------------------------------------------------------
' FR: Restaure les sauts de ligne d'un texte d'aide encodé dans le source VBA.
' EN: Restores line breaks in help text encoded in the VBA source.
'------------------------------------------------------------------------------
Private Function TextCatalogWBS_DecodeTutorialText(ByVal encodedText As String) As String

    TextCatalogWBS_DecodeTutorialText = Replace$(encodedText, "{NL}", vbCrLf)

End Function

