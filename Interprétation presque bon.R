library(shiny)
library(bslib)
library(shinyWidgets)
library(DT)
library(readxl)
library(haven)
library(compareGroups)
library(flextable)
library(officer)
library(writexl)
library(kableExtra)
library(rmarkdown)

# Vérification de Pandoc
if (!rmarkdown::pandoc_available()) {
  stop("Pandoc is required for Word export. Please install it from https://pandoc.org/installing.html")
}

# Fonction pour générer automatiquement les commentaires d'interprétation (optimisée pour études transversales)
generate_analysis_comment <- function(comparegroups_result, var_dependante, var_independante) {
  
  # --- Fonctions utilitaires ---
  format_p <- function(p) {
    if (is.na(p)) return("-")
    if (p < 0.001) return("< 0.001")
    return(format(round(p, 3), nsmall = 3))
  }
  
  format_pr_ci <- function(pr, ci) {
    pr_text <- if (is.na(pr)) "-" else format(round(pr, 2), nsmall = 2)
    ci_text <- if (is.na(ci) || ci == "") "-" else {
      ci_clean <- gsub("[^0-9.;-]", "", ci) # Nettoyer le format
      ci_vals <- as.numeric(unlist(strsplit(ci_clean, ";")))
      if (length(ci_vals) == 2) sprintf("[%.2f;%.2f]", ci_vals[1], ci_vals[2]) else "-"
    }
    list(pr = pr_text, ci = ci_text)
  }
  
  # --- Validation initiale ---
  if (is.null(comparegroups_result) || is.null(comparegroups_result$descr)) {
    return("Aucun résultat statistique détectable.")
  }
  
  # --- Extraction des données ---
  table_data <- tryCatch(as.data.frame(comparegroups_result$descr), error = function(e) NULL)
  if (is.null(table_data) || nrow(table_data) == 0) {
    return("Aucun résultat statistique détectable.")
  }
  
  # --- Identification des colonnes ---
  pr_cols <- c("PR", "Ratio", "RR") # Focus sur PR pour études transversales
  p_val_cols <- c("p.ratio", "p-value", "p.value")
  ci_cols <- c("IC95%", "CI95%", "95%CI", "[95%CI]")
  p_overall_col <- "p.overall"
  
  find_col <- function(patterns, cols) {
    for (p in patterns) {
      match <- grep(p, cols, ignore.case = TRUE, value = TRUE)
      if (length(match) > 0) return(match[1])
    }
    return(NULL)
  }
  
  pr_col_name <- find_col(pr_cols, names(table_data))
  p_val_col_name <- find_col(p_val_cols, names(table_data))
  ci_col_name <- find_col(ci_cols, names(table_data))
  
  # --- Initialisation des commentaires ---
  comments <- list()
  modalite_summaries <- c()
  warnings <- list(missing = NULL, small_n = NULL)
  
  # --- Analyse par modalité ---
  modalites <- rownames(table_data)
  ref_level <- tryCatch(comparegroups_result$ref.y[var_independante], error = function(e) NULL)
  
  for (modalite in modalites) {
    if (grepl("^N$|^%$", modalite, ignore.case = TRUE)) next # Ignorer les lignes N ou %
    if (!is.null(ref_level) && modalite == ref_level) next # Ignorer la modalité de référence
    
    row <- table_data[modalite, ]
    
    # Extraire les valeurs
    pr <- if (!is.null(pr_col_name)) suppressWarnings(as.numeric(row[[pr_col_name]])) else NA
    p_val <- if (!is.null(p_val_col_name)) suppressWarnings(as.numeric(row[[p_val_col_name]])) else NA
    ci <- if (!is.null(ci_col_name)) as.character(row[[ci_col_name]]) else NA
    
    # Formater les valeurs
    formatted <- format_pr_ci(pr, ci)
    pr_text <- formatted$pr
    ci_text <- formatted$ci
    p_text <- format_p(p_val)
    
    comment <- ""
    
    # Détection des effectifs faibles (approximation via pourcentages)
    if (grepl("%", as.character(row[1]), fixed = TRUE)) {
      percent <- suppressWarnings(as.numeric(gsub("[^0-9.]", "", as.character(row[1]))))
      if (!is.na(percent) && percent < 5) {
        warnings$small_n <- paste0("Les effectifs faibles pour la modalité '", modalite, "' de '", var_independante, "' (", percent, "%) limitent la puissance statistique de l’analyse. Les résultats doivent être interprétés avec prudence.")
      }
    }
    
    if (is.na(p_val)) {
      comment <- paste0("La modalité '", modalite, "' de '", var_independante, "' sert de référence pour la comparaison.")
    } else if (p_val < 0.05) {
      if (!is.na(pr) && pr > 1) { # Modèle 1
        comment <- paste0("La modalité '", modalite, "' de '", var_independante, "' est significativement associée à un risque accru de '", var_dependante, "' (PR = ", pr_text, ", IC95% : ", ci_text, ", p = ", p_text, ") dans une étude transversale.")
        modalite_summaries <- c(modalite_summaries, paste0("'", modalite, "' (risque accru)"))
      } else { # Modèle 2
        comment <- paste0("La modalité '", modalite, "' de '", var_independante, "' est associée à un effet protecteur vis-à-vis de '", var_dependante, "' (PR = ", pr_text, ", IC95% : ", ci_text, ", p = ", p_text, ") dans une étude transversale.")
        modalite_summaries <- c(modalite_summaries, paste0("'", modalite, "' (effet protecteur)"))
      }
    } else if (p_val >= 0.05 && p_val < 0.10) { # Modèle 4
      type_effet <- if (!is.na(pr) && pr > 1) "risque accru" else "protecteur"
      comment <- paste0("La modalité '", modalite, "' de '", var_independante, "' suggère une tendance vers un effet ", type_effet, " sur '", var_dependante, "' (PR = ", pr_text, ", IC95% : ", ci_text, ", p = ", p_text, "), mais cette association n’atteint pas le seuil de significativité statistique dans une étude transversale.")
      modalite_summaries <- c(modalite_summaries, paste0("'", modalite, "' (tendance)"))
    } else { # Modèle 3
      comment <- paste0("Aucune association statistiquement significative n’est observée pour la modalité '", modalite, "' de '", var_independante, "' avec '", var_dependante, "' (PR = ", pr_text, ", IC95% : ", ci_text, ", p = ", p_text, "). L’intervalle de confiance inclut la valeur 1, indiquant un effet indéterminé.")
    }
    comments[[modalite]] <- comment
  }
  
  # --- Résumé global (Modèle 5) ---
  p_overall <- if (p_overall_col %in% names(table_data)) suppressWarnings(as.numeric(table_data[1, p_overall_col])) else NA
  p_overall_text <- format_p(p_overall)
  
  global_summary <- ""
  if (!is.na(p_overall)) {
    signif_text <- if (p_overall < 0.05) "significative" else "non significative"
    summary_details <- ""
    if (length(modalite_summaries) > 0) {
      summary_details <- paste0(" Cette association semble principalement portée par les modalités : ", paste(modalite_summaries, collapse = ", "), ".")
    }
    global_summary <- paste0("L’analyse globale indique une association ", signif_text, " entre '", var_independante, "' et '", var_dependante, "' (p.overall = ", p_overall_text, ").", summary_details)
  }
  
  # --- Avertissements (Modèle 6) ---
  missing_info_row <- grep("missing|manquant", rownames(table_data), ignore.case = TRUE)
  if (length(missing_info_row) > 0) {
    missing_text <- table_data[missing_info_row, 1]
    missing_percent <- suppressWarnings(as.numeric(regmatches(missing_text, regexpr("[0-9.]+", missing_text))))
    if (length(missing_percent) > 0 && missing_percent > 5) {
      warnings$missing <- paste0("La proportion de données manquantes (", missing_percent, "%) pour '", var_independante, "' ou '", var_dependante, "' peut limiter la robustesse des résultats. Une analyse complémentaire avec imputation est recommandée.")
    }
  }
  
  # --- Assemblage final ---
  final_comment_parts <- c(
    unlist(comments),
    if (nchar(global_summary) > 0) c("", "--- Résumé ---", global_summary),
    if (!is.null(warnings$missing)) c("", "--- Avertissements ---", warnings$missing),
    if (!is.null(warnings$small_n)) warnings$small_n
  )
  
  if (length(final_comment_parts) == 0) {
    return("Aucun résultat statistique interprétable n'a pu être généré.")
  }
  
  return(paste(final_comment_parts, collapse = "\n"))
}

# Définition du thème personnalisé
custom_theme <- bs_theme(
  version = 5,
  bg = "#ffffff",
  fg = "#2c3e50",
  primary = "#3498db",
  secondary = "#95a5a6",
  success = "#27ae60",
  info = "#17a2b8",
  warning = "#f39c12",
  danger = "#e74c3c",
  base_font = font_google("Inter"),
  heading_font = font_google("Inter", wght = c(400, 600)),
  code_font = font_google("JetBrains Mono")
) %>%
  bs_add_rules("
  .card {
    border: none;
    box-shadow: 0 4px 6px rgba(0, 0, 0, 0.07);
    border-radius: 12px;
    transition: all 0.3s ease;
  }
  .card:hover {
    transform: translateY(-2px);
    box-shadow: 0 8px 25px rgba(0, 0, 0, 0.1);
  }
  .navbar {
    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    border-bottom: 1px solid rgba(0,0,0,0.05);
  }
  .btn {
    border-radius: 8px;
    font-weight: 500;
    transition: all 0.2s ease;
    text-transform: none;
  }
  .btn:hover {
    transform: translateY(-1px);
  }
  .status-success {
    color: #27ae60;
    font-weight: 500;
    display: flex;
    align-items: center;
    gap: 8px;
  }
  .status-error {
    color: #e74c3c;
    font-weight: 500;
  }
  .form-control {
    border-radius: 8px;
    border: 2px solid #e9ecef;
    transition: border-color 0.2s ease;
  }
  .form-control:focus {
    border-color: #3498db;
    box-shadow: 0 0 0 0.2rem rgba(52, 152, 219, 0.25);
  }
  .main-container {
    padding: 2rem;
  }
  .card-header {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    color: white;
    border-radius: 12px 12px 0 0 !important;
    font-weight: 600;
    padding: 1rem 1.5rem;
  }
  .import-section {
    border-left: 4px solid #3498db;
    padding-left: 15px;
    margin-bottom: 20px;
  }
")

# Interface utilisateur améliorée
ui <- page_navbar(
  title = div(
    icon("chart-line", style = "margin-right: 10px;"),
    "Analyse compareGroups - Études transversales",
    style = "font-weight: 600;"
  ),
  theme = custom_theme,
  fillable = TRUE,
  
  # Panel 1: Import des données
  nav_panel(
    title = tagList(icon("upload"), "Import"),
    div(class = "main-container",
        fluidRow(
          column(6,
                 card(
                   card_header(
                     tagList(icon("file-import"), "Importer vos données")
                   ),
                   card_body(
                     div(class = "import-section",
                         h5(tagList(icon("file"), "Depuis un fichier"), style = "color: #3498db; margin-bottom: 15px;"),
                         fileInput("file",
                                   label = div(
                                     icon("cloud-upload", style = "margin-right: 8px;"),
                                     "Sélectionnez votre fichier"
                                   ),
                                   accept = c(".csv", ".xlsx", ".sav", ".dta"),
                                   buttonLabel = "Parcourir",
                                   placeholder = "Aucun fichier sélectionné"),
                         div(
                           icon("info-circle", style = "margin-right: 8px; color: #17a2b8;"),
                           "Formats supportés : CSV, Excel (.xlsx), SPSS (.sav), Stata (.dta)",
                           style = "font-size: 0.9em; color: #6c757d; padding: 10px; background: #f8f9fa; border-radius: 6px;"
                         )
                     ),
                     hr(),
                     div(class = "import-section",
                         h5(tagList(icon("r-project"), "Depuis l'environnement R"), style = "color: #3498db; margin-bottom: 15px;"),
                         selectInput("env_object",
                                     label = div(
                                       icon("database", style = "margin-right: 8px;"),
                                       "Objets disponibles dans l'environnement"
                                     ),
                                     choices = NULL,
                                     selectize = TRUE),
                         actionButton("load_env_object",
                                      tagList(icon("download"), "Charger l'objet sélectionné"),
                                      class = "btn-outline-info w-100",
                                      style = "margin-top: 10px;"),
                         div(
                           icon("lightbulb", style = "margin-right: 8px; color: #f39c12;"),
                           "Seuls les data.frames sont affichés dans la liste",
                           style = "font-size: 0.9em; color: #6c757d; padding: 10px; background: #fff3cd; border-radius: 6px; margin-top: 10px;"
                         )
                     ),
                     br(),
                     uiOutput("file_status_ui")
                   )
                 )
          ),
          column(6,
                 card(
                   card_header(
                     tagList(icon("table"), "Aperçu des données")
                   ),
                   card_body(
                     DTOutput("data_preview")
                   )
                 )
          )
        )
    )
  ),
  
  # Panel 2: Configuration
  nav_panel(
    title = tagList(icon("cogs"), "Configuration"),
    div(class = "main-container",
        fluidRow(
          column(6,
                 card(
                   card_header(
                     tagList(icon("sliders-h"), "Modalités de référence")
                   ),
                   card_body(
                     uiOutput("reference_controls"),
                     br(),
                     actionButton("apply_references",
                                  tagList(icon("check"), "Appliquer les modalités"),
                                  class = "btn-outline-primary w-100",
                                  style = "margin-top: 10px;")
                   )
                 )
          ),
          column(6,
                 card(
                   card_header(
                     tagList(icon("microscope"), "Paramètres d'analyse")
                   ),
                   card_body(
                     selectInput("group_var",
                                 label = tagList(icon("users"), "Variable dépendante"),
                                 choices = NULL),
                     selectizeInput("vars",
                                    label = tagList(icon("list"), "Variables indépendantes"),
                                    choices = NULL,
                                    multiple = TRUE,
                                    options = list(placeholder = "Sélectionnez les variables...")),
                     hr(style = "margin: 1.5rem 0;"),
                     div(class = "row",
                         column(6,
                                materialSwitch("show_ratio",
                                               label = "Afficher PR",
                                               value = TRUE,
                                               status = "primary")
                         ),
                         column(6,
                                materialSwitch("show_ci",
                                               label = "Intervalles de confiance",
                                               value = TRUE,
                                               status = "primary")
                         )
                     ),
                     div(class = "row",
                         column(6,
                                numericInput("max_ylev",
                                             "Max niveaux",
                                             value = 5,
                                             min = 1,
                                             max = 20)
                         )
                     ),
                     br(),
                     actionButton("run_analysis",
                                  tagList(icon("play"), "Lancer l'analyse"),
                                  class = "btn-primary w-100",
                                  style = "height: 50px; font-size: 1.1em; font-weight: 600;")
                   )
                 )
          )
        )
    )
  ),
  
  # Panel 3: Résultats
  nav_panel(
    title = tagList(icon("chart-bar"), "Résultats"),
    div(class = "main-container",
        card(
          card_header(
            tagList(icon("table"), "Tableau descriptif")
          ),
          card_body(
            DTOutput("results_table")
          )
        ),
        br(),
        card(
          card_header(
            tagList(icon("lightbulb"), "Interprétation automatique")
          ),
          card_body(
            div(
              style = "padding: 15px; background-color: #f8f9fa; border-left: 4px solid #007bff; border-radius: 6px;",
              h5("Analyse des associations statistiques (études transversales)",
                 style = "color: #007bff; margin-bottom: 15px;"),
              div(
                id = "interpretation_text",
                style = "line-height: 1.6; color: #2c3e50; font-size: 0.95em;",
                textOutput("auto_interpretation")
              )
            )
          )
        )
    )
  ),
  
  # Panel 4: Export
  nav_panel(
    title = tagList(icon("download"), "Export"),
    div(class = "main-container",
        card(
          card_header(
            tagList(icon("file-word"), "Exporter les résultats")
          ),
          card_body(
            fluidRow(
              column(6,
                     textInput("output_file",
                               label = tagList(icon("file"), "Nom du fichier"),
                               value = "resultats_analyse.docx",
                               placeholder = "nom_fichier.docx")
              ),
              column(6,
                     textInput("report_title",
                               label = tagList(icon("heading"), "Titre du rapport"),
                               value = "Rapport d'Analyse Transversale",
                               placeholder = "Titre de votre rapport")
              )
            ),
            br(),
            div(class = "d-grid",
                downloadButton("export_word",
                               tagList(icon("download"), "Télécharger le rapport Word"),
                               class = "btn-success",
                               style = "height: 50px; font-size: 1.1em; font-weight: 600;")
            ),
            br(),
            uiOutput("export_status_ui")
          )
        )
    )
  )
)

# Serveur
server <- function(input, output, session) {
  # Réactifs pour stocker les données et tableaux
  data <- reactiveVal(NULL)
  original_data <- reactiveVal(NULL)
  all_tables <- reactiveVal(list())
  analysis_result <- reactiveVal(NULL)
  
  # Fonction pour lister les objets data.frame dans l'environnement global
  get_dataframes_from_env <- function() {
    env_objects <- ls(envir = globalenv())
    dataframes <- c()
    
    for (obj_name in env_objects) {
      tryCatch({
        obj <- get(obj_name, envir = globalenv())
        if (is.data.frame(obj)) {
          dataframes <- c(dataframes, obj_name)
        }
      }, error = function(e) {
        # Ignorer les objets qui ne peuvent pas être récupérés
      })
    }
    
    return(dataframes)
  }
  
  # Mise à jour de la liste des objets disponibles
  observe({
    available_dfs <- get_dataframes_from_env()
    choices_list <- setNames(available_dfs,
                             paste0(available_dfs, " (",
                                    sapply(available_dfs, function(x) {
                                      tryCatch({
                                        obj <- get(x, envir = globalenv())
                                        paste(nrow(obj), "lignes,", ncol(obj), "colonnes")
                                      }, error = function(e) "inconnu")
                                    }), ")"))
    
    updateSelectInput(session, "env_object",
                      choices = c("Sélectionnez un objet..." = "", choices_list))
  })
  
  # Charger un objet depuis l'environnement R
  observeEvent(input$load_env_object, {
    req(input$env_object, input$env_object != "")
    
    tryCatch({
      df <- get(input$env_object, envir = globalenv())
      if (!is.data.frame(df)) {
        showNotification("L'objet sélectionné n'est pas un data.frame.", type = "error", duration = 5)
        return()
      }
      
      df[] <- lapply(df, function(x) if (is.character(x)) as.factor(x) else x)
      data(df)
      original_data(df)
      
      updateSelectInput(session, "group_var", choices = c("Aucune" = "", names(df)))
      updateSelectizeInput(session, "vars", choices = names(df))
      
      showNotification(paste("Objet", input$env_object, "chargé avec succès!"),
                       type = "message", duration = 3)
    }, error = function(e) {
      showNotification(paste("Erreur lors du chargement de l'objet :", e$message),
                       type = "error", duration = 5)
    })
  })
  
  # Status visuel pour l'import de fichier
  output$file_status_ui <- renderUI({
    if (!is.null(data())) {
      div(class = "status-success",
          icon("check-circle"),
          paste("Données importées avec succès ! (", nrow(data()), "lignes,", ncol(data()), "colonnes)")
      )
    } else {
      div(class = "status-error", "En attente de données !")
    }
  })
  
  # Status visuel pour l'export
  output$export_status_ui <- renderUI({
    # Cette sortie sera mise à jour dans le downloadHandler
  })
  
  # Importer le fichier
  observeEvent(input$file, {
    req(input$file)
    tryCatch({
      file_ext <- tools::file_ext(input$file$datapath)
      if (file_ext == "csv") {
        df <- read.csv(input$file$datapath, stringsAsFactors = FALSE)
      } else if (file_ext == "xlsx") {
        df <- read_excel(input$file$datapath)
      } else if (file_ext %in% c("sav", "dta")) {
        df <- read_sav(input$file$datapath)
      } else {
        showNotification("Format de fichier non supporté.", type = "error", duration = 5)
        return(NULL)
      }
      
      df[] <- lapply(df, function(x) if (is.character(x)) as.factor(x) else x)
      data(df)
      original_data(df)
      
      updateSelectInput(session, "group_var", choices = c("Aucune" = "", names(df)))
      updateSelectizeInput(session, "vars", choices = names(df))
      
      showNotification("Fichier importé avec succès!", type = "message", duration = 3)
    }, error = function(e) {
      showNotification(paste("Erreur lors de l'importation :", e$message), type = "error", duration = 5)
    })
  })
  
  # Identifier les variables catégorielles
  cat_vars <- reactive({
    req(data())
    names(data())[sapply(data(), is.factor)]
  })
  
  # Génération des contrôles pour les modalités de référence
  output$reference_controls <- renderUI({
    req(cat_vars())
    if (length(cat_vars()) == 0) {
      return(div(
        icon("exclamation-triangle", style = "margin-right: 8px;"),
        "Aucune variable catégorielle détectée.",
        class = "text-warning",
        style = "padding: 15px; background: #fff3cd; border-radius: 6px;"
      ))
    }
    
    lapply(cat_vars(), function(var) {
      div(style = "margin-bottom: 15px;",
          selectInput(
            inputId = paste0("ref_", var),
            label = div(
              icon("tag", style = "margin-right: 6px;"),
              paste("Référence pour", var)
            ),
            choices = levels(data()[[var]]),
            selected = levels(data()[[var]])[1]
          )
      )
    })
  })
  
  # Application des modalités de référence
  observeEvent(input$apply_references, {
    req(original_data())
    tryCatch({
      df <- original_data()
      for (var in cat_vars()) {
        ref_level <- input[[paste0("ref_", var)]]
        if (!is.null(ref_level)) {
          df[[var]] <- relevel(df[[var]], ref = ref_level)
        }
      }
      data(df)
      showNotification("Modalités de référence appliquées!", type = "message", duration = 3)
    }, error = function(e) {
      showNotification(paste("Erreur :", e$message), type = "error", duration = 5)
    })
  })
  
  # Afficher un aperçu des données avec style
  output$data_preview <- renderDT({
    if (is.null(data())) {
      return(datatable(data.frame(Message = "Aucune donnée à afficher"),
                       options = list(dom = 't', ordering = FALSE)))
    }
    
    datatable(head(data(), 10),
              options = list(
                pageLength = 5,
                scrollX = TRUE,
                dom = 'tip',
                language = list(
                  info = "Affichage des 10 premières lignes"
                )
              ),
              class = "table table-striped table-hover")
  })
  
  # Analyse avec compareGroups
  observeEvent(input$run_analysis, {
    req(data(), input$vars)
    
    showNotification("Analyse en cours...", type = "default", duration = 2)
    
    tryCatch({
      if (input$group_var == "") {
        formula <- as.formula(paste("~", paste(input$vars, collapse = "+")))
      } else {
        formula <- as.formula(paste(input$group_var, "~", paste(input$vars, collapse = "+")))
      }
      
      res <- compareGroups(
        formula = formula,
        data = data(),
        max.ylev = input$max_ylev,
        include.label = TRUE,
        compute.ratio = input$show_ratio,
        riskratio = TRUE # Forcer PR pour études transversales
      )
      
      restab <- createTable(
        res,
        show.ratio = input$show_ratio,
        show.ci = input$show_ci,
        show.p.ratio = TRUE,
        digits.p = 3
      )
      
      analysis_result(restab)
      
      current_tables <- all_tables()
      table_title <- if (input$group_var == "") "Tableau descriptif" else paste("Tableau groupé par", input$group_var)
      current_tables[[length(current_tables) + 1]] <- list(
        title = table_title,
        table = res, # Stocker l'objet compareGroups directement
        table_formatted = restab, # Stocker la table formatée
        timestamp = Sys.time()
      )
      all_tables(current_tables)
      
      showNotification("Analyse terminée avec succès!", type = "message", duration = 3)
    }, error = function(e) {
      showNotification(paste("Erreur lors de l'analyse :", e$message), type = "error", duration = 5)
      analysis_result(NULL)
    })
  })
  
  # Génération automatique de l'interprétation
  output$auto_interpretation <- renderText({
    if (is.null(analysis_result())) {
      return("Lancez une analyse pour obtenir une interprétation automatique des résultats.")
    }
    
    # Utiliser group_var comme var_dependante et vars comme var_independante
    var_dependante <- if (input$group_var == "") "la variable étudiée" else input$group_var
    var_independante <- if (length(input$vars) == 0) "les variables sélectionnées" else paste(input$vars, collapse = ", ")
    
    comment <- generate_analysis_comment(analysis_result(), var_dependante, var_independante)
    
    return(comment)
  })
  
  # Afficher le tableau des résultats avec style
  output$results_table <- renderDT({
    if (is.null(analysis_result())) {
      return(datatable(data.frame(Message = "Aucun résultat à afficher. Lancez une analyse."),
                       options = list(dom = 't', ordering = FALSE)))
    }
    
    datatable(as.data.frame(analysis_result()$descr),
              options = list(
                pageLength = 15,
                scrollX = TRUE,
                dom = 'Bfrtip',
                buttons = c('copy', 'csv', 'excel')
              ),
              extensions = 'Buttons',
              class = "table table-striped table-hover")
  })
  
  # Fonction pour créer une flextable de base
  create_flextable <- function(comparegroups_table) {
    table_df <- as.data.frame(comparegroups_table$descr)
    
    if (ncol(table_df) == 0 || nrow(table_df) == 0) {
      return(NULL)
    }
    
    ft <- flextable(table_df) %>%
      theme_vanilla() %>%
      autofit() %>%
      fontsize(size = 9, part = "all") %>%
      align(align = "center", part = "header") %>%
      border_outer(border = fp_border(color = "black", width = 1)) %>%
      border_inner_h(border = fp_border(color = "gray", width = 0.5)) %>%
      border_inner_v(border = fp_border(color = "gray", width = 0.5))
    
    return(ft)
  }
  
  # Fonction pour créer une flextable avec commentaires
  create_identical_flextable_with_comments <- function(comparegroups_table, var_dependante, var_independante) {
    table_df <- as.data.frame(comparegroups_table$descr)
    
    if (ncol(table_df) == 0 || nrow(table_df) == 0) {
      return(list(table = NULL, comment = "Aucune donnée disponible"))
    }
    
    ft <- create_flextable(comparegroups_table)
    
    comment <- generate_analysis_comment(comparegroups_table, var_dependante, var_independante)
    
    return(list(table = ft, comment = comment))
  }
  
  # Export Word
  output$export_word <- downloadHandler(
    filename = function() {
      paste0(tools::file_path_sans_ext(input$output_file), ".docx")
    },
    content = function(file) {
      req(length(all_tables()) > 0)
      
      tryCatch({
        if (length(all_tables()) == 1) {
          table_info <- all_tables()[[1]]
          export2word(table_info$table_formatted, file = file)
        } else {
          doc <- read_docx() %>%
            body_add_par(value = input$report_title, style = "heading 1") %>%
            body_add_par(value = paste("Rapport généré le :", format(Sys.time(), "%d/%m/%Y %H:%M")), style = "Normal") %>%
            body_add_par(value = paste("Nombre de tableaux :", length(all_tables())), style = "Normal") %>%
            body_add_break()
          
          for (i in seq_along(all_tables())) {
            table_info <- all_tables()[[i]]
            var_dependante <- if (input$group_var == "") "la variable étudiée" else input$group_var
            var_independante <- if (length(input$vars) == 0) "les variables sélectionnées" else paste(input$vars, collapse = ", ")
            
            doc <- doc %>%
              body_add_par(value = paste("Tableau", i, ":", table_info$title), style = "heading 2") %>%
              body_add_par(value = paste("Généré le :", format(table_info$timestamp, "%d/%m/%Y %H:%M")), style = "Normal")
            
            ft_with_comments <- create_identical_flextable_with_comments(table_info$table, var_dependante, var_independante)
            
            if (!is.null(ft_with_comments$table)) {
              doc <- doc %>%
                body_add_flextable(value = ft_with_comments$table) %>%
                body_add_par(value = "", style = "Normal") %>%
                body_add_par(value = "Interprétation :", style = "heading 3") %>%
                body_add_par(value = ft_with_comments$comment, style = "Normal")
            } else {
              doc <- doc %>% body_add_par(value = "Erreur : impossible de créer le tableau", style = "Normal")
            }
            
            if (i < length(all_tables())) {
              doc <- doc %>% body_add_break()
            }
          }
          
          print(doc, target = file)
        }
        
        output$export_status_ui <- renderUI({
          div(class = "status-success",
              icon("check-circle"),
              "Fichier Word créé avec succès !")
        })
        showNotification("Exportation réussie !", type = "message", duration = 3)
      }, error = function(e) {
        output$export_status_ui <- renderUI({
          div(class = "status-error",
              icon("exclamation-triangle"),
              paste("Erreur lors de l'exportation :", e$message))
        })
        showNotification(paste("Erreur lors de l'exportation :", e$message), type = "error", duration = 5)
      })
    },
    contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  )
}

# Lancer l'application

shinyApp(ui = ui, server = server)
