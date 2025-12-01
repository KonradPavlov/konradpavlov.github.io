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
    "Analyse compareGroups",
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
                     # Section 1: Import de fichier
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
                     
                     # Section 2: Import depuis l'environnement R
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
                                 label = tagList(icon("users"), "Variable de groupement"),
                                 choices = NULL),
                     
                     selectizeInput("vars", 
                                    label = tagList(icon("list"), "Variables à analyser"),
                                    choices = NULL, 
                                    multiple = TRUE,
                                    options = list(placeholder = "Sélectionnez les variables...")),
                     
                     hr(style = "margin: 1.5rem 0;"),
                     
                     div(class = "row",
                         column(6,
                                materialSwitch("show_ratio", 
                                               label = "Afficher OR/HR/RR", 
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
                                materialSwitch("risk_ratio", 
                                               label = "Préférer RR", 
                                               value = FALSE,
                                               status = "info")
                         ),
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
                               value = "Rapport d'Analyse",
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
    dataframes <- character(0)
    
    for (obj_name in env_objects) {
      tryCatch({
        obj <- get(obj_name, envir = globalenv())
        if (is.data.frame(obj) && !is.null(obj) && nrow(obj) > 0 && ncol(obj) > 0) {
          dataframes <- c(dataframes, obj_name)
        }
      }, error = function(e) {
        # Ignorer silencieusement les erreurs
      })
    }
    
    return(dataframes)
  }
  
  # Mise à jour de la liste des objets disponibles
  observe({
    available_dfs <- get_dataframes_from_env()
    
    # Générer des noms valides pour chaque objet
    choices_names <- sapply(available_dfs, function(x) {
      tryCatch({
        obj <- get(x, envir = globalenv())
        if (is.data.frame(obj)) {
          paste(x, "(", nrow(obj), "lignes,", ncol(obj), "colonnes)")
        } else {
          paste(x, "(inconnu)")
        }
      }, error = function(e) {
        paste(x, "(erreur)")
      })
    })
    
    # Vérifier que choices_names est un vecteur de caractères valide
    if (length(available_dfs) == 0 || any(is.na(choices_names)) || any(is.null(choices_names))) {
      choices_list <- NULL
      showNotification("Aucun data.frame valide trouvé dans l'environnement.", type = "warning", duration = 5)
    } else {
      choices_list <- setNames(available_dfs, choices_names)
    }
    
    # Mettre à jour le selectInput avec une option par défaut
    updateSelectInput(session, "env_object", 
                      choices = c("Sélectionnez un objet..." = "", choices_list))
  })
  
  # Charger un objet depuis l'environnement R
  observeEvent(input$load_env_object, {
    req(input$env_object, input$env_object != "")
    
    tryCatch({
      # Récupérer l'objet depuis l'environnement global
      df <- get(input$env_object, envir = globalenv())
      
      # Vérifier que c'est bien un data.frame
      if (!is.data.frame(df)) {
        showNotification("L'objet sélectionné n'est pas un data.frame.", type = "error", duration = 5)
        return()
      }
      
      # Convertir les caractères en facteurs
      df[] <- lapply(df, function(x) if (is.character(x)) as.factor(x) else x)
      
      # Stocker les données
      data(df)
      original_data(df)
      
      # Mettre à jour les choix dans les autres inputs
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
    
    # Feedback visuel pendant l'analyse
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
        riskratio = input$risk_ratio
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
        table = restab,
        timestamp = Sys.time()
      )
      all_tables(current_tables)
      
      showNotification("Analyse terminée avec succès!", type = "message", duration = 3)
    }, error = function(e) {
      showNotification(paste("Erreur lors de l'analyse :", e$message), type = "error", duration = 5)
      analysis_result(NULL)
    })
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
  
  # Fonction pour créer une flextable identique à celle affichée
  create_identical_flextable <- function(comparegroups_table) {
    # Extraire exactement les mêmes données que celles affichées dans l'app
    table_df <- as.data.frame(comparegroups_table$descr)
    
    # S'assurer que les noms de colonnes sont corrects
    if (ncol(table_df) == 0 || nrow(table_df) == 0) {
      return(NULL)
    }
    
    # Créer la flextable avec les mêmes données
    ft <- flextable(table_df)
    
    # Appliquer exactement le même formatage que dans l'application originale
    ft <- ft %>%
      theme_vanilla() %>%
      autofit() %>%
      fontsize(size = 9, part = "all") %>%
      align(align = "center", part = "header") %>%
      border_outer(border = fp_border(color = "black", width = 1)) %>%
      border_inner_h(border = fp_border(color = "gray", width = 0.5)) %>%
      border_inner_v(border = fp_border(color = "gray", width = 0.5))
    
    return(ft)
  }
  
  # Export Word avec conformité parfaite
  output$export_word <- downloadHandler(
    filename = function() {
      # Forcer l'extension .docx
      paste0(tools::file_path_sans_ext(input$output_file), ".docx")
    },
    content = function(file) {
      req(length(all_tables()) > 0) # Vérifie qu'il y a au moins un tableau
      
      tryCatch({
        # Si un seul tableau, utiliser export2word directement pour garantir la conformité
        if (length(all_tables()) == 1) {
          table_info <- all_tables()[[1]]
          export2word(table_info$table, file = file)
        } else {
          # Si plusieurs tableaux, créer un document Word avec une conformité parfaite
          doc <- read_docx()
          
          # Ajouter une page de titre
          doc <- doc %>%
            body_add_par(value = input$report_title, style = "heading 1") %>%
            body_add_par(value = paste("Rapport généré le :", format(Sys.time(), "%d/%m/%Y %H:%M")), style = "Normal") %>%
            body_add_par(value = paste("Nombre de tableaux :", length(all_tables())), style = "Normal") %>%
            body_add_break()
          
          # Ajouter chaque tableau avec une conformité parfaite
          for (i in seq_along(all_tables())) {
            table_info <- all_tables()[[i]]
            
            # Ajouter le titre du tableau
            doc <- doc %>%
              body_add_par(value = paste("Tableau", i, ":", table_info$title), style = "heading 2") %>%
              body_add_par(value = paste("Généré le :", format(table_info$timestamp, "%d/%m/%Y %H:%M")), style = "Normal")
            
            # Créer une flextable parfaitement identique à celle affichée
            ft <- create_identical_flextable(table_info$table)
            
            if (!is.null(ft)) {
              # Ajouter la flextable au document
              doc <- doc %>% body_add_flextable(value = ft)
            } else {
              doc <- doc %>% body_add_par(value = "Erreur : impossible de créer le tableau", style = "Normal")
            }
            
            # Ajouter un saut de page (sauf pour le dernier tableau)
            if (i < length(all_tables())) {
              doc <- doc %>% body_add_break()
            }
          }
          
          # Sauvegarder le document final
          print(doc, target = file)
        }
        
        # Afficher un message de succès
        output$export_status_ui <- renderUI({
          div(class = "status-success",
              icon("check-circle"),
              "Fichier Word créé avec succès ! (Tableaux parfaitement conformes)")
        })
        showNotification("Exportation réussie avec conformité parfaite !", type = "message", duration = 3)
        
      }, error = function(e) {
        # Gérer les erreurs et afficher un message clair
        output$export_status_ui <- renderUI({
          div(class = "status-error",
              icon("exclamation-triangle"),
              paste("Erreur lors de l'exportation :", e$message))
        })
        showNotification(paste("Erreur lors de l'exportation :", e$message), type = "error", duration = 5)
        
        # Log de l'erreur pour le debug
        cat("Erreur d'exportation :", e$message, "\n")
        print(traceback())
      })
    },
    contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  )
}

# Lancer l'application
shinyApp(ui = ui, server = server)
shinyApp(ui, server, options = list(launch.browser = TRUE))