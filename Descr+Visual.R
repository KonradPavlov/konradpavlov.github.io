# gtsummary Studio Pro - Version Complète avec graphiques Excel-like

library(shiny)
library(shinydashboard)
library(DT)
library(gtsummary)
library(gt)
library(dplyr)
library(readr)
library(readxl)
library(shinyWidgets)
library(shinycssloaders)
library(flextable)
library(officer)
library(questionr)
library(ggplot2)
library(extrafont)  # Pour Times New Roman
library(plotly)     # Pour l'interactivité Plotly

theme_gtsummary_language(language = "fr", decimal.mark = ",", big.mark = " ", set_theme = TRUE)

# Fonctions utilitaires
detect_var_type <- function(x) {
  if (is.null(x) || length(x) == 0) return("unknown")
  x_clean <- x[!is.na(x)]
  if (length(x_clean) == 0) return("unknown")
  if (is.numeric(x_clean)) {
    unique_vals <- length(unique(x_clean))
    if (unique_vals <= 10 && all(x_clean %% 1 == 0, na.rm = TRUE)) return("categorical")
    return("continuous")
  } else if (is.factor(x_clean) || is.character(x_clean)) {
    return("categorical")
  } else if (is.logical(x_clean)) {
    return("binary")
  }
  return("unknown")
}

suggest_grouping_var <- function(data) {
  cat_vars <- names(data)[sapply(data, function(x) detect_var_type(x) %in% c("categorical", "binary"))]
  if (length(cat_vars) > 0) {
    levels_count <- sapply(cat_vars, function(v) length(unique(data[[v]][!is.na(data[[v]])])))
    valid_vars <- cat_vars[levels_count >= 2 & levels_count <= 10]
    if (length(valid_vars) > 0) return(valid_vars[1])
    return(cat_vars[1])
  }
  return("")
}

get_available_dataframes <- function() {
  all_objects <- ls(envir = .GlobalEnv)
  dataframes <- character(0)
  for (obj_name in all_objects) {
    obj <- get(obj_name, envir = .GlobalEnv)
    if (is.data.frame(obj) && nrow(obj) > 0) dataframes <- c(dataframes, obj_name)
  }
  return(dataframes)
}

parse_recodage <- function(expr) {
  pairs <- strsplit(expr, ";")[[1]]
  recodes <- list()
  for (pair in pairs) {
    parts <- strsplit(pair, "=")[[1]]
    if (length(parts) == 2) recodes[[trimws(parts[2])]] <- trimws(parts[1])
  }
  return(recodes)
}

# Interface utilisateur
ui <- dashboardPage(
  dashboardHeader(title = "gtsummary Studio Pro", titleWidth = 280),
  dashboardSidebar(
    width = 280,
    sidebarMenu(
      menuItem("📊 Données", tabName = "data", icon = icon("database")),
      menuItem("🎯 Modalités", tabName = "reference", icon = icon("cog")),
      menuItem("📈 Descriptif", tabName = "descriptive", icon = icon("table")),
      menuItem("🔍 Croisé", tabName = "crosstab", icon = icon("th")),
      menuItem("📊 Visualisations", tabName = "visualizations", icon = icon("chart-bar")),
      menuItem("📥 Export", tabName = "export", icon = icon("download"))
    )
  ),
  dashboardBody(
    tags$head(tags$style(HTML("
      .content-wrapper { background-color: #f8f9fa; }
      .alert-info { background-color: #d1ecf1; border-color: #bee5eb; color: #0c5460; padding: 10px; border-radius: 5px; }
      .error-message { color: red; font-weight: bold; }
      .success-message { color: green; font-weight: bold; }
      .reference-section { background-color: #f8f9fa; padding: 15px; border-radius: 5px; margin-bottom: 10px; }
      .excel-theme .ggplot { font-family: 'Calibri', sans-serif; }
      .excel-theme .title { font-size: 16pt; font-weight: bold; }
      .excel-theme .axis-text { font-size: 11pt; }
      .preview-item { border: 1px solid #ddd; padding: 10px; margin-bottom: 10px; border-radius: 5px; }
      .preview-thumbnail { max-width: 200px; max-height: 150px; }
      .data-table { font-family: 'Times New Roman', serif; font-size: 12pt; }
    "))),
    tabItems(
      # Onglet Données
      tabItem(tabName = "data",
              fluidRow(
                box(title = "📂 Import des Données", status = "primary", solidHeader = TRUE, width = 6,
                    radioButtons("data_source", "Source des données:",
                                 choices = list("Fichier CSV" = "csv", 
                                                "Fichier Excel" = "excel",
                                                "Données R (environnement)" = "r_env"),
                                 selected = "csv"),
                    conditionalPanel("input.data_source == 'csv'",
                                     fileInput("csv_file", "Sélectionner un fichier CSV:",
                                               accept = ".csv")),
                    conditionalPanel("input.data_source == 'excel'",
                                     fileInput("excel_file", "Sélectionner un fichier Excel:",
                                               accept = c(".xlsx", ".xls"))),
                    conditionalPanel("input.data_source == 'r_env'",
                                     div(class = "alert-info",
                                         "💡 Sélectionnez un data.frame déjà chargé dans votre environnement R."),
                                     br(),
                                     selectInput("r_dataframe", "Choisir un data.frame:",
                                                 choices = NULL,
                                                 width = "100%"),
                                     br(),
                                     actionButton("refresh_dataframes", "🔄 Actualiser la liste", 
                                                  class = "btn-info btn-sm")),
                    actionButton("load_data", "Charger les données", 
                                 class = "btn-primary", style = "margin-top: 10px;")
                ),
                box(title = "🔍 Informations", status = "info", solidHeader = TRUE, width = 6,
                    htmlOutput("data_info")
                )
              ),
              fluidRow(
                box(title = "👁️ Aperçu des Données", status = "success", solidHeader = TRUE, width = 12,
                    withSpinner(DT::dataTableOutput("data_preview"))
                )
              )
      ),
      # Onglet Modalités
      tabItem(tabName = "reference",
              fluidRow(
                box(title = "🎯 Gestion des Modalités de Référence", status = "primary", solidHeader = TRUE, width = 12,
                    div(class = "alert-info",
                        "💡 Définissez les modalités de référence pour vos variables catégorielles."),
                    br(),
                    conditionalPanel("output.data_loaded",
                                     uiOutput("reference_controls"),
                                     br(),
                                     actionButton("apply_references", "✅ Appliquer les modalités de référence", 
                                                  class = "btn-success"),
                                     br(), br(),
                                     htmlOutput("reference_status")
                    ),
                    conditionalPanel("!output.data_loaded",
                                     p("Veuillez charger des données dans l'onglet 'Données'.", 
                                       class = "error-message"))
                )
              )
      ),
      # Onglet Descriptif
      tabItem(tabName = "descriptive",
              fluidRow(
                column(4,
                       box(title = "⚙️ Configuration", status = "primary", solidHeader = TRUE, width = 12,
                           selectInput("desc_variables", "Variables à analyser:", 
                                       choices = NULL, multiple = TRUE),
                           checkboxInput("use_grouping", "Grouper par une variable", FALSE),
                           conditionalPanel("input.use_grouping == true",
                                            selectInput("desc_by", "Variable de groupement:", 
                                                        choices = NULL)),
                           checkboxInput("desc_missing", "Afficher les valeurs manquantes", TRUE),
                           checkboxInput("desc_overall", "Ajouter une colonne 'Total'", FALSE),
                           checkboxInput("add_p_values", "Ajouter les tests statistiques", FALSE),
                           numericInput("decimal_places", "Nombre de décimales:", 
                                        value = 1, min = 0, max = 3),
                           br(),
                           actionButton("generate_desc", "🚀 Générer le tableau", 
                                        class = "btn-success btn-block")
                       )
                ),
                column(8,
                       box(title = "📊 Tableau Descriptif", status = "success", solidHeader = TRUE, width = 12,
                           withSpinner(htmlOutput("descriptive_table")),
                           htmlOutput("desc_messages")
                       )
                )
              )
      ),
      # Onglet Croisé
      tabItem(tabName = "crosstab",
              fluidRow(
                column(4,
                       box(title = "⚙️ Configuration", status = "primary", solidHeader = TRUE, width = 12,
                           selectInput("cross_row", "Variable en ligne:", choices = NULL),
                           selectInput("cross_col", "Variable en colonne:", choices = NULL),
                           checkboxInput("cross_percent", "Afficher les pourcentages", TRUE),
                           p("Type de pourcentage : Ligne"),
                           checkboxInput("add_chi2", "Test du Chi²", TRUE),
                           br(),
                           actionButton("generate_cross", "🚀 Générer le tableau", 
                                        class = "btn-success btn-block")
                       )
                ),
                column(8,
                       box(title = "🔍 Tableau Croisé", status = "success", solidHeader = TRUE, width = 12,
                           withSpinner(htmlOutput("crosstab_table")),
                           htmlOutput("cross_messages")
                       )
                )
              )
      ),
      # Onglet Visualisations
      tabItem(tabName = "visualizations",
              fluidRow(
                # Panneau de configuration
                column(4,
                       box(title = "⚙️ Configuration du Graphique", status = "primary", 
                           solidHeader = TRUE, width = 12,
                           
                           # Sélection variable principale
                           selectInput("plot_var", "Variable à visualiser :", 
                                       choices = NULL, width = "100%"),
                           
                           # Variable de regroupement
                           selectInput("group_var", "Variable de regroupement (facultatif) :", 
                                       choices = c("Aucun" = ""), width = "100%"),
                           
                           # Type de graphique (sera mis à jour automatiquement)
                           uiOutput("plot_type_ui"),
                           
                           # Recommandations intelligentes
                           uiOutput("plot_recommendations"),
                           
                           # Personnalisation
                           h5("🎨 Personnalisation"),
                           textInput("plot_title", "Titre du graphique :", 
                                     value = "Graphique univarié", width = "100%"),
                           textInput("plot_xlab", "Label de l'axe X :", 
                                     value = "", width = "100%"),
                           textInput("plot_ylab", "Label de l'axe Y :", 
                                     value = "", width = "100%"),
                           
                           # Options couleurs
                           radioButtons("color_palette", "Palette de couleurs :", 
                                        choices = c("Palette Excel" = "excel", 
                                                    "Couleur unie" = "single",
                                                    "Palette moderne" = "modern"),
                                        selected = "excel", inline = TRUE),
                           
                           # Type d'axe Y pour les graphiques en barres
                           conditionalPanel(
                             condition = "input.plot_type == 'bar'",
                             radioButtons("y_axis_type", "Type d'axe Y :",
                                          choices = c("Effectifs" = "counts", 
                                                      "Pourcentages" = "percent"),
                                          selected = "counts", inline = TRUE)
                           ),
                           
                           # Interactivité
                           checkboxInput("use_plotly", "Graphique interactif (Plotly)", 
                                         value = FALSE),
                           
                           # Boutons d'action
                           br(),
                           actionButton("generate_plot", "🚀 Générer le graphique", 
                                        class = "btn-success btn-block"),
                           br(),
                           actionButton("add_to_export", "➕ Ajouter à l'export", 
                                        class = "btn-info btn-block")
                       )
                ),
                
                # Panneau d'affichage
                column(8,
                       box(title = "📊 Graphique Généré", status = "success", 
                           solidHeader = TRUE, width = 12,
                           
                           # Zone d'affichage du graphique
                           uiOutput("plot_display_ui"),
                           
                           # Informations sur le graphique
                           uiOutput("plot_info"),
                           
                           # Export rapide
                           br(),
                           fluidRow(
                             column(6,
                                    textInput("plot_filename", "Nom du fichier :", 
                                              value = "graphique", width = "100%")),
                             column(3,
                                    selectInput("plot_format", "Format :", 
                                                choices = c("PNG" = "png", "JPG" = "jpg", 
                                                            "PDF" = "pdf", "SVG" = "svg"),
                                                selected = "png", width = "100%")),
                             column(3,
                                    br(),
                                    downloadButton("download_plot", "💾 Télécharger", 
                                                   class = "btn-warning btn-block"))
                           )
                       )
                )
              ),
              
              # Liste des graphiques générés
              fluidRow(
                box(title = "📋 Graphiques Générés", status = "info", 
                    solidHeader = TRUE, width = 12,
                    uiOutput("visualization_preview_list")
                )
              )
      ),
      # Onglet Export
      tabItem(tabName = "export",
              fluidRow(
                box(title = "📥 Export des Tableaux et Graphiques", status = "primary", solidHeader = TRUE, width = 12,
                    div(class = "alert-info",
                        "💡 Sélectionnez les éléments à exporter après prévisualisation."),
                    br(),
                    h4("📋 Liste des éléments générés :"),
                    uiOutput("export_preview_list"),
                    br(),
                    h4("🔍 Prévisualisation :"),
                    uiOutput("preview_content"),
                    br(),
                    fluidRow(
                      column(3,
                             downloadButton("download_html", "📄 Télécharger HTML", 
                                            class = "btn-info btn-block")),
                      column(3,
                             downloadButton("download_csv", "📊 Télécharger CSV", 
                                            class = "btn-warning btn-block")),
                      column(3,
                             downloadButton("download_word", "📝 Télécharger Word", 
                                            class = "btn-success btn-block")),
                      column(3,
                             textInput("report_title", "Titre du rapport:", 
                                       value = "Analyse Statistique"))
                    )
                )
              )
      )
    )
  )
)

# Serveur
server <- function(input, output, session) {
  current_data <- reactiveVal(NULL)
  original_data <- reactiveVal(NULL)
  current_table <- reactiveVal(NULL)
  all_tables <- reactiveVal(list())
  reference_levels <- reactiveVal(list())
  current_plot <- reactiveVal(NULL)
  all_plots <- reactiveVal(list())
  plot_counter <- reactiveVal(0)
  selected_preview <- reactiveVal(NULL)
  
  output$data_loaded <- reactive({ !is.null(current_data()) })
  outputOptions(output, "data_loaded", suspendWhenHidden = FALSE)
  
  update_dataframe_choices <- function() {
    available_dfs <- get_available_dataframes()
    if (length(available_dfs) > 0) {
      choices <- setNames(available_dfs, paste0(available_dfs, " (", sapply(available_dfs, function(x) {
        df <- get(x, envir = .GlobalEnv)
        paste0(nrow(df), " obs, ", ncol(df), " vars")
      }), ")"))
      updateSelectInput(session, "r_dataframe", choices = choices)
    } else {
      updateSelectInput(session, "r_dataframe", choices = list("Aucun data.frame disponible" = ""))
    }
  }
  
  observe({ if (input$data_source == "r_env") update_dataframe_choices() })
  observeEvent(input$refresh_dataframes, {
    update_dataframe_choices()
    showNotification("Liste des data.frames actualisée !", type = "message")
  })
  
  observeEvent(input$load_data, {
    tryCatch({
      if (input$data_source == "csv" && !is.null(input$csv_file)) {
        data <- read_csv(input$csv_file$datapath, show_col_types = FALSE)
        original_data(data)
        current_data(data)
        showNotification("Fichier CSV chargé avec succès!", type = "message")
      } else if (input$data_source == "excel" && !is.null(input$excel_file)) {
        data <- read_excel(input$excel_file$datapath)
        original_data(data)
        current_data(data)
        showNotification("Fichier Excel chargé avec succès!", type = "message")
      } else if (input$data_source == "r_env" && !is.null(input$r_dataframe) && input$r_dataframe != "") {
        if (exists(input$r_dataframe, envir = .GlobalEnv)) {
          data <- get(input$r_dataframe, envir = .GlobalEnv)
          if (!is.data.frame(data)) {
            showNotification("L'objet sélectionné n'est pas un data.frame valide.", type = "error")
            return()
          }
          if (nrow(data) == 0) {
            showNotification("Le data.frame sélectionné est vide.", type = "warning")
            return()
          }
          original_data(data)
          current_data(data)
          showNotification(paste("Data.frame", input$r_dataframe, "chargé avec succès!"), type = "message")
        } else {
          showNotification("Le data.frame sélectionné n'existe plus.", type = "error")
          update_dataframe_choices()
          return()
        }
      } else {
        showNotification("Veuillez sélectionner une source de données valide.", type = "warning")
        return()
      }
      all_tables(list())
      reference_levels(list())
    }, error = function(e) {
      showNotification(paste("Erreur lors du chargement :", e$message), type = "error")
    })
  })
  
  observe({
    req(current_data())
    data <- current_data()
    all_vars <- names(data)
    cat_vars <- all_vars[sapply(data, function(x) is.factor(x) || is.character(x) || is.logical(x))]
    
    updateSelectInput(session, "plot_var", choices = all_vars)
    updateSelectInput(session, "desc_variables", choices = all_vars, selected = all_vars[1:min(5, length(all_vars))])
    updateSelectInput(session, "desc_by", choices = c("Aucun" = "", setNames(cat_vars, cat_vars)))
    updateSelectInput(session, "cross_row", choices = cat_vars)
    updateSelectInput(session, "cross_col", choices = cat_vars)
    updateSelectInput(session, "group_var", 
                      choices = c("Aucun" = "", setNames(cat_vars, cat_vars)))
  })
  
  output$data_preview <- DT::renderDataTable({
    req(current_data())
    datatable(current_data(), options = list(scrollX = TRUE), class = "data-table")
  })
  
  output$data_info <- renderUI({
    req(current_data())
    data <- current_data()
    HTML(paste0(
      "<p><strong>Nombre d'observations :</strong> ", nrow(data), "</p>",
      "<p><strong>Nombre de variables :</strong> ", ncol(data), "</p>"
    ))
  })
  
  # ----------------------------------------------
  # Partie Serveur : Gestion de l'onglet Visualisations
  # ----------------------------------------------
  # Mise à jour des choix de variables pour les visualisations
  observe({
    req(current_data())
    
    tryCatch({
      data <- current_data()
      if (is.null(data) || nrow(data) == 0) return()
      
      all_vars <- names(data)
      
      # Variables catégorielles pour le regroupement
      cat_vars <- all_vars[sapply(data, function(x) {
        var_type <- detect_var_type(x)
        var_type %in% c("categorical", "binary")
      })]
      
      updateSelectInput(session, "plot_var", choices = all_vars)
      updateSelectInput(session, "group_var", 
                        choices = c("Aucun" = "", setNames(cat_vars, cat_vars)))
    }, error = function(e) {
      # Gestion silencieuse des erreurs
    })
  })
  
  # Interface dynamique pour sélectionner le type de graphique
  output$plot_type_ui <- renderUI({
    req(current_data(), input$plot_var)
    
    tryCatch({
      data <- current_data()
      var <- input$plot_var
      
      if (is.null(data) || !var %in% names(data)) {
        return(div(class = "alert alert-danger", "Variable non trouvée"))
      }
      
      # Nettoyer les données pour l'analyse
      var_data <- data[[var]]
      var_data_clean <- var_data[!is.na(var_data)]
      
      if (length(var_data_clean) == 0) {
        return(div(class = "alert alert-warning", "Aucune donnée valide pour cette variable"))
      }
      
      var_type <- detect_var_type(var_data_clean)
      unique_levels <- length(unique(var_data_clean))
      
      # Choix selon le type de variable
      if (var_type == "continuous") {
        choices <- c("Histogramme" = "histogram", 
                     "Boîte à moustaches" = "boxplot",
                     "Diagramme en barres" = "bar")
        default <- "histogram"
      } else if (var_type %in% c("categorical", "binary")) {
        if (unique_levels == 2) {
          choices <- c("Camembert" = "pie", 
                       "Diagramme en barres" = "bar")
          default <- "pie"
        } else if (unique_levels <= 20) {
          choices <- c("Diagramme en barres" = "bar", 
                       "Camembert" = "pie")
          default <- "bar"
        } else {
          choices <- c("Diagramme en barres" = "bar")
          default <- "bar"
        }
      } else {
        choices <- c("Diagramme en barres" = "bar")
        default <- "bar"
      }
      
      radioButtons("plot_type", "Type de graphique :", 
                   choices = choices, selected = default, inline = TRUE)
      
    }, error = function(e) {
      div(class = "alert alert-danger", "Erreur lors de la détection du type de variable")
    })
  })
  
  # Recommandations intelligentes pour le choix du graphique
  output$plot_recommendations <- renderUI({
    req(current_data(), input$plot_var)
    
    tryCatch({
      data <- current_data()
      var <- input$plot_var
      
      if (is.null(data) || !var %in% names(data)) return(NULL)
      
      var_data <- data[[var]]
      var_data_clean <- var_data[!is.na(var_data)]
      
      if (length(var_data_clean) == 0) return(NULL)
      
      var_type <- detect_var_type(var_data_clean)
      unique_levels <- length(unique(var_data_clean))
      missing_count <- sum(is.na(var_data))
      
      # Générer recommandations
      if (var_type == "continuous") {
        recommendation <- "📊 Recommandé : Histogramme pour les variables quantitatives"
        icon_color <- "#5B9BD5"
      } else if (var_type %in% c("categorical", "binary")) {
        if (unique_levels == 2) {
          recommendation <- "🥧 Recommandé : Camembert pour les variables dichotomiques"
          icon_color <- "#70AD47"
        } else if (unique_levels <= 10) {
          recommendation <- "📊 Recommandé : Diagramme en barres pour les variables catégorielles"
          icon_color <- "#FFC000"
        } else {
          recommendation <- "⚠️ Attention : Trop de modalités pour une visualisation claire"
          icon_color <- "#FF6B6B"
        }
      } else {
        recommendation <- "⚠️ Type de variable non reconnu"
        icon_color <- "#FF6B6B"
      }
      
      # Ajouter info sur les valeurs manquantes
      if (missing_count > 0) {
        missing_info <- paste0(" (", missing_count, " valeurs manquantes)")
      } else {
        missing_info <- ""
      }
      
      div(
        style = paste0("background-color: ", icon_color, "20; border-left: 4px solid ", 
                       icon_color, "; padding: 10px; margin: 10px 0; border-radius: 4px;"),
        p(paste0(recommendation, missing_info), 
          style = paste0("color: ", icon_color, "; font-weight: bold; margin: 0;"))
      )
      
    }, error = function(e) { NULL })
  })
  
  # Génération du graphique
  observeEvent(input$generate_plot, {
    req(input$plot_var, input$plot_type, current_data())
    
    tryCatch({
      data <- current_data()
      var <- input$plot_var
      plot_type <- input$plot_type
      
      # Vérifications
      if (is.null(data) || !var %in% names(data)) {
        showNotification("Variable non trouvée", type = "error")
        return()
      }
      
      if (plot_type == "none") {
        showNotification("Type de graphique non disponible", type = "error")
        return()
      }
      
      # Nettoyer les données
      data_clean <- data[!is.na(data[[var]]), ]
      if (nrow(data_clean) == 0) {
        showNotification("Aucune donnée valide pour cette variable", type = "error")
        return()
      }
      
      # Variable de regroupement
      group_var <- NULL
      if (!is.null(input$group_var) && input$group_var != "" && input$group_var %in% names(data_clean)) {
        group_var <- input$group_var
        data_clean <- data_clean[!is.na(data_clean[[group_var]]), ]
        if (nrow(data_clean) == 0) {
          showNotification("Aucune donnée valide après regroupement", type = "error")
          return()
        }
      }
      
      # Titres et labels
      title <- if (is.null(input$plot_title) || input$plot_title == "") paste("Graphique de", var) else input$plot_title
      xlab <- if (is.null(input$plot_xlab) || input$plot_xlab == "") var else input$plot_xlab
      ylab <- if (is.null(input$plot_ylab) || input$plot_ylab == "") {
        if (input$plot_type == "bar" && !is.null(input$y_axis_type)) {
          if (input$y_axis_type == "percent") "Pourcentage (%)" else "Effectifs"
        } else "Fréquence"
      } else input$plot_ylab
      
      # Thème professionnel Excel-like avec inclinaison des étiquettes
      excel_theme <- theme_minimal() +
        theme(
          text = element_text(family = "Times New Roman", colour = "#333333", size = 12),
          plot.title = element_text(size = 16, face = "bold", hjust = 0.5, margin = margin(b = 20)),
          axis.title = element_text(size = 14, face = "bold"),
          axis.text = element_text(size = 12, color = "#333333"),
          axis.text.x = element_text(angle = 45, hjust = 1, vjust = 1), # Inclinaison à 45°
          panel.grid.major = element_line(color = "#E0E0E0", size = 0.5),
          panel.grid.minor = element_blank(),
          plot.background = element_rect(fill = "#FFFFFF", color = NA),
          panel.background = element_rect(fill = "#FFFFFF", color = NA),
          plot.margin = margin(20, 20, 20, 20),
          legend.position = "right",
          legend.title = element_text(size = 12, face = "bold"),
          legend.text = element_text(size = 11),
          strip.text = element_text(size = 12, face = "bold")
        )
      
      # Palettes de couleurs
      excel_colors <- c("#5B9BD5", "#70AD47", "#FFC000", "#FF6B6B", "#A855F7", 
                        "#06B6D4", "#E74C3C", "#9B59B6", "#F39C12", "#1ABC9C")
      modern_colors <- c("#667eea", "#764ba2", "#f093fb", "#f5576c", "#4facfe", 
                         "#43e97b", "#fa709a", "#fee140", "#a8edea", "#d299c2")
      single_color <- "#5B9BD5"
      
      # Sélection des couleurs
      colors <- switch(input$color_palette,
                       "excel" = excel_colors,
                       "modern" = modern_colors,
                       "single" = single_color,
                       excel_colors)
      
      # Génération du graphique selon le type
      if (plot_type == "histogram") {
        n_bins <- max(10, min(30, ceiling(sqrt(nrow(data_clean)))))
        
        p <- ggplot(data_clean, aes_string(x = var)) + 
          geom_histogram(fill = colors[1], color = "#FFFFFF", bins = n_bins, alpha = 0.8) + 
          labs(title = title, x = xlab, y = ylab) +
          excel_theme
        
      } else if (plot_type == "boxplot") {
        if (is.null(group_var)) {
          p <- ggplot(data_clean, aes_string(x = "''", y = var)) + 
            geom_boxplot(fill = colors[1], color = "#333333", alpha = 0.8, width = 0.5) + 
            labs(title = title, x = "", y = ylab) +
            excel_theme +
            theme(axis.text.x = element_blank(), axis.ticks.x = element_blank())
        } else {
          n_groups <- length(unique(data_clean[[group_var]]))
          group_colors <- if (input$color_palette == "single") {
            rep(single_color, n_groups)
          } else {
            colors[1:min(n_groups, length(colors))]
          }
          
          p <- ggplot(data_clean, aes_string(x = group_var, y = var, fill = group_var)) + 
            geom_boxplot(color = "#333333", alpha = 0.8) + 
            scale_fill_manual(values = group_colors) +
            labs(title = title, x = xlab, y = ylab, fill = group_var) +
            excel_theme
        }
        
      } else if (plot_type == "bar") {
        var_type <- detect_var_type(data_clean[[var]])
        
        if (var_type == "continuous") {
          # Découper en intervalles pour les variables continues
          n_breaks <- min(10, length(unique(data_clean[[var]])))
          breaks <- pretty(data_clean[[var]], n = n_breaks)
          data_clean$intervals <- cut(data_clean[[var]], breaks = breaks, include.lowest = TRUE)
          
          if (is.null(group_var)) {
            freq_data <- as.data.frame(table(data_clean$intervals))
            names(freq_data) <- c(var, "Value")
          } else {
            freq_data <- as.data.frame(table(data_clean$intervals, data_clean[[group_var]]))
            names(freq_data) <- c(var, group_var, "Value")
          }
        } else {
          # Variables catégorielles
          if (is.null(group_var)) {
            freq_data <- as.data.frame(table(data_clean[[var]]))
            names(freq_data) <- c(var, "Value")
          } else {
            freq_data <- as.data.frame(table(data_clean[[var]], data_clean[[group_var]]))
            names(freq_data) <- c(var, group_var, "Value")
          }
        }
        
        # Conversion en pourcentages si demandé
        if (!is.null(input$y_axis_type) && input$y_axis_type == "percent") {
          if (is.null(group_var)) {
            freq_data$Value <- round(freq_data$Value / sum(freq_data$Value) * 100, 1)
          } else {
            freq_data <- freq_data %>%
              group_by(!!sym(group_var)) %>%
              mutate(Value = round(Value / sum(Value) * 100, 1)) %>%
              ungroup()
          }
        }
        
        freq_data[[var]] <- as.character(freq_data[[var]])
        
        if (is.null(group_var)) {
          n_bars <- nrow(freq_data)
          bar_colors <- if (input$color_palette == "single") {
            rep(single_color, n_bars)
          } else {
            colors[1:min(n_bars, length(colors))]
          }
          
          p <- ggplot(freq_data, aes_string(x = var, y = "Value")) + 
            geom_col(fill = bar_colors, color = "#FFFFFF", width = 0.7, alpha = 0.8) +
            labs(title = title, x = xlab, y = ylab) +
            excel_theme +
            geom_text(aes(label = Value), vjust = -0.5, size = 4, 
                      family = "Times New Roman", color = "#333333")
        } else {
          n_groups <- length(unique(freq_data[[group_var]]))
          group_colors <- if (input$color_palette == "single") {
            rep(single_color, n_groups)
          } else {
            colors[1:min(n_groups, length(colors))]
          }
          
          p <- ggplot(freq_data, aes_string(x = var, y = "Value", fill = group_var)) + 
            geom_col(position = position_dodge(width = 0.8), color = "#FFFFFF", width = 0.7, alpha = 0.8) +
            scale_fill_manual(values = group_colors) +
            labs(title = title, x = xlab, y = ylab, fill = group_var) +
            excel_theme +
            geom_text(aes(label = Value, group = !!sym(group_var)), 
                      position = position_dodge(width = 0.8), vjust = -0.5, 
                      size = 4, family = "Times New Roman", color = "#333333")
        }
        
      } else if (plot_type == "pie") {
        pie_data <- as.data.frame(table(data_clean[[var]]))
        names(pie_data) <- c(var, "Value")
        pie_data$Percentage <- round(pie_data$Value / sum(pie_data$Value) * 100, 1)
        pie_data$Label <- paste0(pie_data[[var]], "\n", pie_data$Percentage, "%")
        
        n_slices <- nrow(pie_data)
        pie_colors <- if (input$color_palette == "single") {
          rep(single_color, n_slices)
        } else {
          colors[1:min(n_slices, length(colors))]
        }
        
        p <- ggplot(pie_data, aes_string(x = "''", y = "Value", fill = var)) + 
          geom_col(width = 1, color = "#FFFFFF", size = 1) + 
          coord_polar(theta = "y", start = 0) + 
          scale_fill_manual(values = pie_colors) +
          labs(title = title, fill = var) +
          excel_theme +
          theme(
            axis.title = element_blank(), 
            axis.text = element_blank(), 
            axis.ticks = element_blank(), 
            panel.grid = element_blank(), 
            axis.line = element_blank()
          ) +
          geom_text(aes(label = Label), position = position_stack(vjust = 0.5), 
                    size = 4, family = "Times New Roman", color = "#FFFFFF")
      }
      
      # Appliquer Plotly si demandé
      if (!is.null(input$use_plotly) && input$use_plotly && plot_type != "pie") {
        p <- ggplotly(p, tooltip = c("x", "y", "fill")) %>%
          layout(
            font = list(family = "Times New Roman", size = 12),
            title = list(text = title, font = list(size = 16, family = "Times New Roman")),
            showlegend = !is.null(group_var)
          )
      }
      
      # Sauvegarder le graphique
      current_plot(p)
      
      # Incrémenter le compteur
      plot_counter(plot_counter() + 1)
      
      # Ajouter aux graphiques sauvegardés
      current_plots <- all_plots()
      new_plot <- list(
        id = plot_counter(),
        title = title,
        type = plot_type,
        variable = var,
        group_var = group_var,
        plot = p,
        is_plotly = !is.null(input$use_plotly) && input$use_plotly && plot_type != "pie",
        timestamp = Sys.time()
      )
      current_plots[[length(current_plots) + 1]] <- new_plot
      all_plots(current_plots)
      
      showNotification("Graphique généré avec succès !", type = "success")
      
    }, error = function(e) {
      showNotification(paste("Erreur lors de la génération :", e$message), type = "error")
    })
  })
  
  # Interface d'affichage du graphique
  output$plot_display_ui <- renderUI({
    req(current_plot())
    
    if (inherits(current_plot(), "plotly")) {
      plotlyOutput("current_plot_plotly", height = "500px")
    } else {
      plotOutput("current_plot_static", height = "500px")
    }
  })
  
  # Affichage du graphique statique
  output$current_plot_static <- renderPlot({
    req(current_plot())
    if (!inherits(current_plot(), "plotly")) {
      current_plot()
    }
  }, height = 500)
  
  # Affichage du graphique interactif
  output$current_plot_plotly <- renderPlotly({
    req(current_plot())
    if (inherits(current_plot(), "plotly")) {
      current_plot()
    }
  })
  
  # Informations sur le graphique
  output$plot_info <- renderUI({
    req(current_plot())
    
    plots <- all_plots()
    if (length(plots) > 0) {
      latest_plot <- plots[[length(plots)]]
      
      HTML(paste0(
        "<div class='alert alert-info' style='margin-top: 15px;'>",
        "<strong>📊 Informations du graphique :</strong><br>",
        "• Variable : ", latest_plot$variable, "<br>",
        if (!is.null(latest_plot$group_var)) paste0("• Regroupement : ", latest_plot$group_var, "<br>"),
        "• Type : ", latest_plot$type, "<br>",
        "• Interactif : ", if (latest_plot$is_plotly) "Oui" else "Non", "<br>",
        "• Généré le : ", format(latest_plot$timestamp, "%d/%m/%Y %H:%M"),
        "</div>"
      ))
    }
  })
  
  # Téléchargement du graphique
  output$download_plot <- downloadHandler(
    filename = function() {
      filename <- if (is.null(input$plot_filename) || input$plot_filename == "") "graphique" else input$plot_filename
      format <- if (is.null(input$plot_format)) "png" else input$plot_format
      paste0(filename, ".", format)
    },
    content = function(file) {
      req(current_plot())
      
      tryCatch({
        if (inherits(current_plot(), "plotly")) {
          # Export Plotly
          if (input$plot_format %in% c("png", "jpg", "jpeg")) {
            if (requireNamespace("webshot", quietly = TRUE)) {
              plotly::export(current_plot(), file = file, format = input$plot_format, 
                             width = 1200, height = 800)
            } else {
              showNotification("Package webshot requis pour exporter les graphiques Plotly", type = "error")
              return()
            }
          } else {
            showNotification("Format non supporté pour les graphiques interactifs", type = "error")
            return()
          }
        } else {
          # Export ggplot2
          width <- if (input$plot_format == "pdf") 12 else 1200
          height <- if (input$plot_format == "pdf") 8 else 800
          dpi <- 300
          
          ggsave(file, plot = current_plot(), device = input$plot_format, 
                 width = width, height = height, dpi = dpi, 
                 units = if (input$plot_format == "pdf") "in" else "px", 
                 bg = "white")
        }
        
        showNotification("Graphique téléchargé avec succès !", type = "success")
      }, error = function(e) {
        showNotification(paste("Erreur lors du téléchargement :", e$message), type = "error")
      })
    }
  )
  
  # Ajouter à l'export
  observeEvent(input$add_to_export, {
    req(current_plot())
    
    tryCatch({
      plots <- all_plots()
      if (length(plots) > 0 && !is.null(current_plot())) {
        showNotification("Graphique ajouté à la liste d'export !", type = "success")
      } else {
        showNotification("Aucun graphique à ajouter ou erreur lors de l'ajout", type = "warning")
      }
    }, error = function(e) {
      showNotification(paste("Erreur lors de l'ajout à l'export :", e$message), type = "error")
    })
  })
  
  # Liste des graphiques générés
  output$visualization_preview_list <- renderUI({
    plots <- all_plots()
    
    if (length(plots) == 0) {
      return(div(class = "alert alert-info", "Aucun graphique généré pour le moment."))
    }
    
    preview_items <- lapply(seq_along(plots), function(i) {
      plot_info <- plots[[i]]
      
      div(class = "preview-item", style = "border: 1px solid #ddd; padding: 15px; margin-bottom: 10px; border-radius: 5px; background-color: #f9f9f9;",
          fluidRow(
            column(3,
                   h5(paste("Graphique", i)),
                   p(strong("Variable: "), plot_info$variable),
                   if (!is.null(plot_info$group_var)) p(strong("Regroupement: "), plot_info$group_var),
                   p(strong("Type: "), plot_info$type),
                   p(strong("Créé le: "), format(plot_info$timestamp, "%d/%m/%Y %H:%M"))
            ),
            column(6,
                   h5(plot_info$title),
                   p(style = "font-style: italic;", 
                     if (plot_info$is_plotly) "Graphique interactif" else "Graphique statique")
            ),
            column(3,
                   actionButton(paste0("preview_", i), "👁️ Prévisualiser", 
                                class = "btn-info btn-sm", style = "margin-bottom: 5px;"),
                   br(),
                   actionButton(paste0("delete_", i), "🗑️ Supprimer", 
                                class = "btn-danger btn-sm")
            )
          )
      )
    })
    
    do.call(tagList, preview_items)
  })
  
  # Gestion des boutons de prévisualisation et suppression
  observeEvent(all_plots(), {
    plots <- all_plots()
    
    if (length(plots) > 0) {
      lapply(seq_along(plots), function(i) {
        # Prévisualisation
        observeEvent(input[[paste0("preview_", i)]], {
          if (i <= length(all_plots())) {
            current_plot(all_plots()[[i]]$plot)
            showNotification(paste("Prévisualisation du graphique", i), type = "message")
          }
        }, ignoreInit = TRUE)
        
        # Suppression
        observeEvent(input[[paste0("delete_", i)]], {
          if (i <= length(all_plots())) {
            current_plots <- all_plots()
            current_plots[[i]] <- NULL
            all_plots(current_plots)
            showNotification(paste("Graphique", i, "supprimé"), type = "warning")
          }
        }, ignoreInit = TRUE)
      })
    }
  })
  
  output$reference_controls <- renderUI({
    req(current_data())
    data <- current_data()
    cat_vars <- names(data)[sapply(data, function(x) is.factor(x) || is.character(x) || is.logical(x))]
    if (length(cat_vars) == 0) return(p("Aucune variable catégorielle détectée.", class = "error-message"))
    controls <- lapply(cat_vars, function(var) {
      levels <- unique(data[[var]][!is.na(data[[var]])])
      if (length(levels) > 1) {
        current_ref <- reference_levels()[[var]]
        if (is.null(current_ref) || !current_ref %in% levels) current_ref <- levels[1]
        div(class = "reference-section",
            h5(paste("📌", var)),
            selectInput(
              inputId = paste0("ref_", var),
              label = "Modalité de référence :",
              choices = levels,
              selected = current_ref,
              width = "100%"
            )
        )
      }
    })
    controls[!sapply(controls, is.null)]
  })
  
  observeEvent(input$apply_references, {
    req(original_data())
    tryCatch({
      data <- original_data()
      new_refs <- list()
      cat_vars <- names(data)[sapply(data, function(x) is.factor(x) || is.character(x) || is.logical(x))]
      for (var in cat_vars) {
        ref_input_id <- paste0("ref_", var)
        if (!is.null(input[[ref_input_id]])) {
          new_refs[[var]] <- input[[ref_input_id]]
          if (is.character(data[[var]])) data[[var]] <- factor(data[[var]])
          if (is.factor(data[[var]])) data[[var]] <- relevel(data[[var]], ref = input[[ref_input_id]])
        }
      }
      current_data(data)
      reference_levels(new_refs)
      output$reference_status <- renderUI({
        HTML(paste0(
          "<div class='success-message'>",
          "✅ Modalités de référence appliquées avec succès !<br>",
          "<strong>Variables modifiées:</strong> ", paste(names(new_refs), collapse = ", "),
          "</div>"
        ))
      })
      showNotification("Modalités de référence appliquées !", type = "message")
    }, error = function(e) {
      output$reference_status <- renderUI({
        HTML(paste0("<p class='error-message'>❌ Erreur: ", e$message, "</p>"))
      })
    })
  })
  
  observeEvent(input$generate_desc, {
    req(current_data(), input$desc_variables)
    output$desc_messages <- renderUI({ NULL })
    tryCatch({
      data <- current_data()
      selected_vars <- input$desc_variables
      missing_vars <- selected_vars[!selected_vars %in% names(data)]
      if (length(missing_vars) > 0) {
        output$desc_messages <- renderUI({
          HTML(paste("<p class='error-message'>Variables non trouvées:", paste(missing_vars, collapse = ", "), "</p>"))
        })
        return()
      }
      data_subset <- data[, selected_vars, drop = FALSE]
      if (input$use_grouping && !is.null(input$desc_by) && input$desc_by != "" && input$desc_by %in% names(data)) {
        data_subset[[input$desc_by]] <- data[[input$desc_by]]
        tbl <- data_subset %>%
          tbl_summary(
            by = all_of(input$desc_by),
            missing = if(input$desc_missing) "always" else "no",
            digits = all_continuous() ~ input$decimal_places
          ) %>%
          bold_labels()
        if (input$desc_overall) tbl <- tbl %>% add_overall(last = TRUE)
        if (input$add_p_values) tbl <- tbl %>% add_p() %>% bold_p()
        table_title <- paste("Tableau descriptif groupé par", input$desc_by)
      } else {
        tbl <- data_subset %>%
          tbl_summary(
            missing = if(input$desc_missing) "always" else "no",
            digits = all_continuous() ~ input$decimal_places
          ) %>%
          bold_labels()
        table_title <- "Tableau descriptif"
      }
      current_table(tbl)
      add_table_to_list(tbl, table_title, "Descriptif")
      output$descriptive_table <- renderUI({
        tbl %>% as_gt() %>% as_raw_html()
      })
      output$desc_messages <- renderUI({
        HTML("<p class='success-message'>✅ Tableau généré avec succès!</p>")
      })
    }, error = function(e) {
      output$desc_messages <- renderUI({
        HTML(paste("<p class='error-message'>❌ Erreur:", e$message, "</p>"))
      })
    })
  })
  
  observeEvent(input$generate_cross, {
    req(current_data(), input$cross_row, input$cross_col)
    output$cross_messages <- renderUI({ NULL })
    tryCatch({
      data <- current_data()
      if (!input$cross_row %in% names(data) || !input$cross_col %in% names(data)) {
        output$cross_messages <- renderUI({
          HTML("<p class='error-message'>❌ Variables sélectionnées non valides.</p>")
        })
        return()
      }
      tbl <- data %>%
        tbl_cross(
          row = all_of(input$cross_row),
          col = all_of(input$cross_col),
          percent = if(input$cross_percent) "row" else "none"
        ) %>%
        bold_labels()
      if (input$add_chi2) tbl <- tbl %>% add_p() %>% bold_p()
      current_table(tbl)
      table_title <- paste("Tableau croisé:", input$cross_row, "x", input$cross_col)
      add_table_to_list(tbl, table_title, "Croisé")
      output$crosstab_table <- renderUI({
        tbl %>% as_gt() %>% as_raw_html()
      })
      output$cross_messages <- renderUI({
        HTML("<p class='success-message'>✅ Tableau croisé généré avec succès!</p>")
      })
    }, error = function(e) {
      output$cross_messages <- renderUI({
        HTML(paste("<p class='error-message'>❌ Erreur:", e$message, "</p>"))
      })
    })
  })
  
  add_table_to_list <- function(table, title, type) {
    current_tables <- all_tables()
    new_table <- list(
      title = title,
      type = type,
      table = table,
      timestamp = Sys.time()
    )
    current_tables[[length(current_tables) + 1]] <- new_table
    all_tables(current_tables)
  }
  
  output$tables_summary <- renderUI({
    tables <- all_tables()
    if (length(tables) == 0) return(p("Aucun tableau généré pour le moment.", class = "error-message"))
    summary_html <- lapply(seq_along(tables), function(i) {
      table_info <- tables[[i]]
      paste0(
        "<div style='margin-bottom: 10px; padding: 10px; background-color: #f8f9fa; border-radius: 5px;'>",
        "<strong>", i, ". ", table_info$title, "</strong><br>",
        "<small>Type: ", table_info$type, " | Généré le: ", 
        format(table_info$timestamp, "%d/%m/%Y %H:%M"), "</small>",
        "</div>"
      )
    })
    HTML(paste(summary_html, collapse = ""))
  })
  
  output$export_preview_list <- renderUI({
    tables <- all_tables()
    plots <- all_plots()
    all_items <- c(tables, plots)
    
    if (length(all_items) == 0) {
      return(div(class = "alert alert-info", "Aucun élément généré pour le moment."))
    }
    
    preview_items <- lapply(seq_along(all_items), function(i) {
      item <- all_items[[i]]
      is_table <- i <= length(tables)
      
      div(class = "preview-item", style = "border: 1px solid #ddd; padding: 15px; margin-bottom: 10px; border-radius: 5px; background-color: #f9f9f9;",
          fluidRow(
            column(8,
                   h5(if (is_table) paste("Tableau", i) else paste("Graphique", i - length(tables))),
                   p(strong("Titre: "), item$title),
                   p(strong("Type: "), if (is_table) item$type else item$type),
                   p(strong("Créé le: "), format(item$timestamp, "%d/%m/%Y %H:%M"))
            ),
            column(4,
                   checkboxInput(paste0("export_item_", i), "Inclure dans l'export", value = TRUE)
            )
          )
      )
    })
    
    do.call(tagList, preview_items)
  })
  
  output$preview_content <- renderUI({
    selected <- selected_preview()
    if (is.null(selected)) {
      return(div(class = "alert alert-info", "Sélectionnez un élément pour une prévisualisation."))
    }
    
    if (selected$type %in% c("Descriptif", "Croisé")) {
      HTML(selected$table %>% as_gt() %>% as_raw_html())
    } else if (selected$type %in% c("histogram", "boxplot", "bar", "pie")) {
      if (inherits(selected$plot, "plotly")) {
        plotlyOutput("preview_plotly", height = "400px")
      } else {
        plotOutput("preview_static", height = "400px")
      }
    }
  })
  
  output$preview_plotly <- renderPlotly({
    req(selected_preview())
    selected <- selected_preview()
    if (inherits(selected$plot, "plotly")) selected$plot
  })
  
  output$preview_static <- renderPlot({
    req(selected_preview())
    selected <- selected_preview()
    if (!inherits(selected$plot, "plotly")) selected$plot
  })
  
  observeEvent(input$download_html, {
    tables <- all_tables()
    plots <- all_plots()
    all_items <- c(tables, plots)
    selected_items <- sapply(seq_along(all_items), function(i) {
      input[[paste0("export_item_", i)]] %||% FALSE
    })
    
    if (!any(selected_items)) {
      showNotification("Aucun élément sélectionné pour l'exportation.", type = "warning")
      return()
    }
    
    output$download_html <- downloadHandler(
      filename = function() { paste0("rapport_", Sys.Date(), ".html") },
      content = function(file) {
        html_content <- c(
          "<!DOCTYPE html>",
          "<html><head>",
          "<title>", input$report_title, "</title>",
          "<meta charset='UTF-8'>",
          "<style>body { font-family: 'Times New Roman', serif; margin: 20px; } img { max-width: 100%; }</style>",
          "</head><body>",
          "<h1>", input$report_title, "</h1>",
          "<p>Généré le: ", Sys.Date(), "</p>"
        )
        index <- 1
        for (i in seq_along(tables)) {
          if (selected_items[index]) {
            table_info <- tables[[i]]
            html_content <- c(html_content,
                              "<h2>", paste("Tableau ", i, ": ", table_info$title), "</h2>",
                              table_info$table %>% as_gt() %>% as_raw_html())
          }
          index <- index + 1
        }
        for (i in seq_along(plots)) {
          if (selected_items[index]) {
            plot_info <- plots[[i]]
            temp_file <- tempfile(fileext = ".png")
            if (plot_info$is_plotly) {
              plotly::export(plot_info$plot, file = temp_file, format = "png")
            } else {
              ggsave(temp_file, plot = plot_info$plot, device = "png", width = 12, height = 8, dpi = 300)
            }
            img_data <- base64enc::base64encode(readBin(temp_file, "raw", file.info(temp_file)$size))
            html_content <- c(html_content,
                              "<h2>", paste("Graphique ", i, ": ", plot_info$title), "</h2>",
                              "<img src='data:image/png;base64,", img_data, "' alt='", plot_info$title, "'>")
            unlink(temp_file)
          }
          index <- index + 1
        }
        html_content <- c(html_content, "</body></html>")
        writeLines(html_content, file)
        showNotification("Rapport HTML généré avec succès !", type = "success")
      }
    )
  })
  
  output$download_csv <- downloadHandler(
    filename = function() { paste0("donnees_", Sys.Date(), ".csv") },
    content = function(file) {
      req(current_data())
      write_csv(current_data(), file)
      showNotification("Fichier CSV téléchargé avec succès !", type = "success")
    }
  )
  
  output$download_word <- downloadHandler(
    filename = function() { paste0("rapport_complet_", Sys.Date(), ".docx") },
    content = function(file) {
      tables <- all_tables()
      plots <- all_plots()
      all_items <- c(tables, plots)
      selected_items <- sapply(seq_along(all_items), function(i) {
        input[[paste0("export_item_", i)]] %||% FALSE
      })
      
      if (!any(selected_items)) {
        showNotification("Aucun élément sélectionné pour l'exportation.", type = "warning")
        return()
      }
      
      tryCatch({
        doc <- read_docx()
        doc <- doc %>%
          body_add_par(value = input$report_title, style = "heading 1") %>%
          body_add_par(value = paste("Rapport généré le:", format(Sys.Date(), "%d/%m/%Y")), style = "Normal") %>%
          body_add_par(value = paste("Nombre d'éléments:", sum(selected_items)), style = "Normal") %>%
          body_add_break()
        index <- 1
        for (i in seq_along(tables)) {
          if (selected_items[index]) {
            table_info <- tables[[i]]
            doc <- doc %>%
              body_add_par(value = paste("Tableau", i, ":", table_info$title), style = "heading 2") %>%
              body_add_par(value = paste("Type:", table_info$type, "| Généré le:", format(table_info$timestamp, "%d/%m/%Y %H:%M")), style = "Normal") %>%
              body_add_flextable(table_info$table %>% as_flex_table()) %>%
              body_add_break()
          }
          index <- index + 1
        }
        for (i in seq_along(plots)) {
          if (selected_items[index]) {
            plot_info <- plots[[i]]
            temp_file <- tempfile(fileext = ".png")
            if (plot_info$is_plotly) {
              plotly::export(plot_info$plot, file = temp_file, format = "png")
            } else {
              ggsave(temp_file, plot = plot_info$plot, device = "png", width = 12, height = 8, dpi = 300)
            }
            doc <- doc %>%
              body_add_par(value = paste("Graphique", i, ": ", plot_info$title), style = "heading 2") %>%
              body_add_par(value = paste("Type:", plot_info$type, "| Généré le:", format(plot_info$timestamp, "%d/%m/%Y %H:%M")), style = "Normal") %>%
              body_add_img(src = temp_file, width = 6, height = 4) %>%
              body_add_break()
            unlink(temp_file)
          }
          index <- index + 1
        }
        print(doc, target = file)
        showNotification("Rapport Word généré avec succès !", type = "message")
      }, error = function(e) {
        showNotification(paste("Erreur lors de la génération du rapport :", e$message), type = "error")
      })
    }
  )
}

# Lancement de l'application
shinyApp(ui, server)