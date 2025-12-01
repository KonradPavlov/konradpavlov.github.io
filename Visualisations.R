# Visualisations - Application Shiny pour l'analyse et la visualisation de données

library(shiny)
library(shinydashboard)
library(DT)
library(dplyr)
library(readr)
library(readxl)
library(shinyWidgets)
library(shinycssloaders)
library(ggplot2)
library(plotly)
library(viridis)
library(RColorBrewer)
library(colourpicker)
library(extrafont)

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

get_available_dataframes <- function() {
  all_objects <- ls(envir = .GlobalEnv)
  dataframes <- character(0)
  for (obj_name in all_objects) {
    obj <- get(obj_name, envir = .GlobalEnv)
    if (is.data.frame(obj) && nrow(obj) > 0) dataframes <- c(dataframes, obj_name)
  }
  return(dataframes)
}

# Interface utilisateur
ui <- dashboardPage(
  dashboardHeader(title = "Visualisations", titleWidth = 280),
  dashboardSidebar(
    width = 280,
    sidebarMenu(
      menuItem("📊 Données", tabName = "data", icon = icon("database")),
      menuItem("🎯 Modalités", tabName = "reference", icon = icon("cocktail")),
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
      # Onglet Visualisations
      tabItem(tabName = "visualizations",
              fluidRow(
                column(4,
                       box(title = "⚙️ Configuration Avancée du Graphique", 
                           status = "primary", solidHeader = TRUE, width = 12,
                           h4("📊 Variables et Type de Graphique", 
                              style = "color: #337ab7; border-bottom: 2px solid #337ab7; padding-bottom: 5px;"),
                           selectInput("plot_var", "Variable principale à visualiser :", 
                                       choices = NULL, width = "100%"),
                           selectInput("group_var", "Variable de regroupement (optionnel) :", 
                                       choices = c("Aucun" = ""), width = "100%"),
                           uiOutput("plot_type_ui"),
                           uiOutput("plot_recommendations"),
                           hr(),
                           h4("🏷️ Titres et Labels", 
                              style = "color: #5cb85c; border-bottom: 2px solid #5cb85c; padding-bottom: 5px;"),
                           textInput("plot_title", "Titre du graphique :", 
                                     value = "Graphique univarié", width = "100%"),
                           fluidRow(
                             column(6, textInput("plot_xlab", "Label axe X :", value = "", width = "100%")),
                             column(6, textInput("plot_ylab", "Label axe Y :", value = "", width = "100%"))
                           ),
                           textInput("plot_subtitle", "Sous-titre (optionnel) :", value = "", width = "100%"),
                           textInput("plot_caption", "Note de bas de page (optionnel) :", value = "", width = "100%"),
                           hr(),
                           h4("🔤 Personnalisation des Polices", 
                              style = "color: #f0ad4e; border-bottom: 2px solid #f0ad4e; padding-bottom: 5px;"),
                           fluidRow(
                             column(6, 
                                    selectInput("font_family", "Famille de police :",
                                                choices = c("Arial" = "Arial", 
                                                            "Times New Roman" = "Times New Roman",
                                                            "Calibri" = "Calibri",
                                                            "Helvetica" = "Helvetica",
                                                            "Georgia" = "Georgia",
                                                            "Verdana" = "Verdana"),
                                                selected = "Arial", width = "100%")),
                             column(6, 
                                    numericInput("base_font_size", "Taille de base :", 
                                                 value = 12, min = 8, max = 24, step = 1, width = "100%"))
                           ),
                           fluidRow(
                             column(6,
                                    h5("📊 Titre Principal"),
                                    numericInput("title_size", "Taille :", value = 16, min = 10, max = 30, step = 1),
                                    selectInput("title_style", "Style :",
                                                choices = c("Normal" = "plain", "Gras" = "bold", 
                                                            "Italique" = "italic", "Gras-Italique" = "bold.italic"),
                                                selected = "bold")),
                             column(6,
                                    h5("🏷️ Labels des Axes"),
                                    numericInput("axis_title_size", "Taille :", value = 14, min = 8, max = 20, step = 1),
                                    selectInput("axis_title_style", "Style :",
                                                choices = c("Normal" = "plain", "Gras" = "bold", 
                                                            "Italique" = "italic", "Gras-Italique" = "bold.italic"),
                                                selected = "bold"))
                           ),
                           fluidRow(
                             column(6,
                                    h5("📈 Texte des Axes"),
                                    numericInput("axis_text_size", "Taille :", value = 12, min = 8, max = 16, step = 1),
                                    selectInput("axis_text_style", "Style :",
                                                choices = c("Normal" = "plain", "Gras" = "bold", 
                                                            "Italique" = "italic", "Gras-Italique" = "bold.italic"),
                                                selected = "plain")),
                             column(6,
                                    h5("🏷️ Légende"),
                                    numericInput("legend_text_size", "Taille :", value = 11, min = 8, max = 16, step = 1),
                                    selectInput("legend_title_style", "Style titre :",
                                                choices = c("Normal" = "plain", "Gras" = "bold", 
                                                            "Italique" = "italic", "Gras-Italique" = "bold.italic"),
                                                selected = "bold"))
                           ),
                           hr(),
                           h4("🎨 Personnalisation des Couleurs", 
                              style = "color: #d9534f; border-bottom: 2px solid #d9534f; padding-bottom: 5px;"),
                           radioButtons("color_mode", "Mode de couleur :",
                                        choices = c("Palette prédéfinie" = "palette",
                                                    "Couleurs personnalisées" = "custom",
                                                    "Couleur unique" = "single"),
                                        selected = "palette", inline = TRUE),
                           conditionalPanel(
                             condition = "input.color_mode == 'palette'",
                             selectInput("color_palette", "Palette :",
                                         choices = c("Excel moderne" = "excel",
                                                     "Viridis" = "viridis",
                                                     "Set1" = "set1",
                                                     "Set2" = "set2",
                                                     "Dark2" = "dark2",
                                                     "Paired" = "paired",
                                                     "Pastel1" = "pastel1",
                                                     "Accent" = "accent"),
                                         selected = "excel", width = "100%")
                           ),
                           conditionalPanel(
                             condition = "input.color_mode == 'single'",
                             colourInput("single_color", "Couleur principale :", value = "#5B9BD5")
                           ),
                           conditionalPanel(
                             condition = "input.color_mode == 'custom'",
                             numericInput("n_colors", "Nombre de couleurs :", value = 5, min = 2, max = 10, step = 1),
                             uiOutput("custom_colors_ui")
                           ),
                           hr(),
                           h4("📏 Dimensions et Mise en Page", 
                              style = "color: #5bc0de; border-bottom: 2px solid #5bc0de; padding-bottom: 5px;"),
                           fluidRow(
                             column(6,
                                    numericInput("plot_width", "Largeur (px) :", value = 800, min = 400, max = 1600, step = 50)),
                             column(6,
                                    numericInput("plot_height", "Hauteur (px) :", value = 600, min = 300, max = 1200, step = 50))
                           ),
                           selectInput("legend_position", "Position de la légende :",
                                       choices = c("Droite" = "right", "Gauche" = "left", 
                                                   "Haut" = "top", "Bas" = "bottom", "Masquer" = "none"),
                                       selected = "right", width = "100%"),
                           h5("📐 Marges (en pixels)"),
                           fluidRow(
                             column(3, numericInput("margin_top", "Haut :", value = 20, min = 0, max = 100, step = 5)),
                             column(3, numericInput("margin_right", "Droite :", value = 20, min = 0, max = 100, step = 5)),
                             column(3, numericInput("margin_bottom", "Bas :", value = 20, min = 0, max = 100, step = 5)),
                             column(3, numericInput("margin_left", "Gauche :", value = 20, min = 0, max = 100, step = 5))
                           ),
                           hr(),
                           h4("⚙️ Options Avancées", 
                              style = "color: #337ab7; border-bottom: 2px solid #337ab7; padding-bottom: 5px;"),
                           conditionalPanel(
                             condition = "input.plot_type == 'bar'",
                             radioButtons("y_axis_type", "Type d'axe Y :",
                                          choices = c("Effectifs" = "counts", "Pourcentages" = "percent"),
                                          selected = "counts", inline = TRUE)
                           ),
                           conditionalPanel(
                             condition = "input.plot_type == 'bar'",
                             numericInput("x_axis_rotation", "Rotation labels X (degrés) :", 
                                          value = 0, min = -90, max = 90, step = 15)
                           ),
                           checkboxInput("show_grid_major", "Afficher grille principale", value = TRUE),
                           checkboxInput("show_grid_minor", "Afficher grille secondaire", value = FALSE),
                           checkboxInput("use_plotly", "Graphique interactif (Plotly)", value = FALSE),
                           selectInput("background_theme", "Thème de fond :",
                                       choices = c("Blanc" = "white", "Gris clair" = "light_gray", 
                                                   "Sombre" = "dark", "Minimal" = "minimal"),
                                       selected = "white", width = "100%"),
                           hr(),
                           h4("🚀 Actions", 
                              style = "color: #5cb85c; border-bottom: 2px solid #5cb85c; padding-bottom: 5px;"),
                           actionButton("generate_plot", "🎨 Générer le graphique", 
                                        class = "btn-success btn-block", 
                                        style = "margin-bottom: 10px; font-weight: bold;"),
                           actionButton("preview_settings", "👁️ Prévisualiser les paramètres", 
                                        class = "btn-info btn-block", 
                                        style = "margin-bottom: 10px;"),
                           actionButton("reset_settings", "🔄 Réinitialiser", 
                                        class = "btn-warning btn-block", 
                                        style = "margin-bottom: 10px;"),
                           actionButton("add_to_export", "💾 Ajouter à l'export", 
                                        class = "btn-primary btn-block")
                       )
                ),
                column(8,
                       box(title = "📊 Graphique Généré", status = "success", 
                           solidHeader = TRUE, width = 12,
                           uiOutput("plot_display_container"),
                           uiOutput("plot_detailed_info"),
                           hr(),
                           h4("💾 Export Avancé", 
                              style = "color: #f0ad4e; border-bottom: 2px solid #f0ad4e; padding-bottom: 5px;"),
                           fluidRow(
                             column(4,
                                    textInput("plot_filename", "Nom du fichier :", 
                                              value = "graphique_personnalise", width = "100%")),
                             column(4,
                                    selectInput("plot_format", "Format :",
                                                choices = c("PNG (Haute qualité)" = "png",
                                                            "JPG (Compact)" = "jpg",
                                                            "PDF (Vectoriel)" = "pdf",
                                                            "SVG (Web)" = "svg",
                                                            "EPS (Publication)" = "eps"),
                                                selected = "png", width = "100%")),
                             column(4,
                                    numericInput("export_dpi", "DPI :", value = 300, min = 72, max = 600, step = 72, width = "100%"))
                           ),
                           fluidRow(
                             column(6,
                                    numericInput("export_width", "Largeur export (pouces) :", 
                                                 value = 10, min = 2, max = 20, step = 0.5)),
                             column(6,
                                    numericInput("export_height", "Hauteur export (pouces) :", 
                                                 value = 8, min = 2, max = 16, step = 0.5))
                           ),
                           fluidRow(
                             column(6,
                                    downloadButton("download_plot", "📥 Télécharger", 
                                                   class = "btn-warning btn-block")),
                             column(6,
                                    actionButton("copy_code", "📋 Copier le code R", 
                                                 class = "btn-info btn-block"))
                           )
                       )
                )
              ),
              fluidRow(
                box(title = "📋 Historique des Graphiques", status = "info", 
                    solidHeader = TRUE, width = 12,
                    fluidRow(
                      column(4,
                             selectInput("filter_plot_type", "Filtrer par type :",
                                         choices = c("Tous" = "all"), selected = "all", width = "100%")),
                      column(4,
                             selectInput("filter_variable", "Filtrer par variable :",
                                         choices = c("Toutes" = "all"), selected = "all", width = "100%")),
                      column(4,
                             actionButton("clear_history", "🗑️ Vider l'historique", 
                                          class = "btn-danger", style = "margin-top: 25px;"))
                    ),
                    hr(),
                    uiOutput("enhanced_visualization_list")
                )
              )
      ),
      # Onglet Export
      tabItem(tabName = "export",
              fluidRow(
                box(title = "📥 Export des Graphiques", status = "primary", solidHeader = TRUE, width = 12,
                    div(class = "alert-info",
                        "💡 Sélectionnez les graphiques à exporter après prévisualisation."),
                    br(),
                    h4("📋 Liste des graphiques générés :"),
                    uiOutput("export_preview_list"),
                    br(),
                    h4("🔍 Prévisualisation :"),
                    uiOutput("preview_content"),
                    br(),
                    fluidRow(
                      column(4,
                             downloadButton("download_html", "📄 Télécharger HTML", 
                                            class = "btn-info btn-block")),
                      column(4,
                             downloadButton("download_csv", "📊 Télécharger CSV", 
                                            class = "btn-warning btn-block")),
                      column(4,
                             textInput("report_title", "Titre du rapport:", 
                                       value = "Analyse Visuelle"))
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
  reference_levels <- reactiveVal(list())
  current_plot <- reactiveVal(NULL)
  plot_history <- reactiveVal(list())
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
  
  output$custom_colors_ui <- renderUI({
    req(input$n_colors)
    color_inputs <- lapply(1:input$n_colors, function(i) {
      default_colors <- c("#5B9BD5", "#70AD47", "#FFC000", "#FF6B6B", "#A855F7", 
                          "#06B6D4", "#E74C3C", "#9B59B6", "#F39C12", "#1ABC9C")
      fluidRow(
        column(6, h6(paste("Couleur", i, ":"))),
        column(6, colourInput(paste0("custom_color_", i), NULL, 
                              value = default_colors[i %% length(default_colors) + 1], 
                              showColour = "background"))
      )
    })
    do.call(tagList, color_inputs)
  })
  
  output$plot_display_container <- renderUI({
    req(current_plot())
    plot_width <- if (!is.null(input$plot_width)) paste0(input$plot_width, "px") else "800px"
    plot_height <- if (!is.null(input$plot_height)) paste0(input$plot_height, "px") else "600px"
    tags$div(
      style = paste0("width: 100%; max-width: ", plot_width, "; margin: 0 auto;"),
      if (inherits(current_plot(), "plotly")) {
        plotlyOutput("current_plot_plotly", width = "100%", height = plot_height)
      } else {
        plotOutput("current_plot_static", width = "100%", height = plot_height)
      }
    )
  })
  
  observe({
    req(current_data())
    tryCatch({
      data <- current_data()
      if (is.null(data) || nrow(data) == 0) return()
      all_vars <- names(data)
      cat_vars <- all_vars[sapply(data, function(x) {
        var_type <- detect_var_type(x)
        var_type %in% c("categorical", "binary")
      })]
      updateSelectInput(session, "plot_var", choices = all_vars)
      updateSelectInput(session, "group_var", 
                        choices = c("Aucun" = "", setNames(cat_vars, cat_vars)))
      updateSelectInput(session, "filter_variable", 
                        choices = c("Toutes" = "all", setNames(all_vars, all_vars)))
    }, error = function(e) {
      showNotification("Erreur lors de la mise à jour des variables", type = "error")
    })
  })
  
  output$plot_type_ui <- renderUI({
    req(current_data(), input$plot_var)
    tryCatch({
      data <- current_data()
      var <- input$plot_var
      if (is.null(data) || !var %in% names(data)) {
        return(div(class = "alert alert-danger", "⚠️ Variable non trouvée dans les données"))
      }
      var_data <- data[[var]]
      var_data_clean <- var_data[!is.na(var_data)]
      if (length(var_data_clean) == 0) {
        return(div(class = "alert alert-warning", "⚠️ Aucune donnée valide pour cette variable"))
      }
      var_type <- detect_var_type(var_data_clean)
      unique_levels <- length(unique(var_data_clean))
      n_observations <- length(var_data_clean)
      if (var_type == "continuous") {
        choices <- c("Histogramme" = "histogram", 
                     "Densité" = "density",
                     "Boîte à moustaches" = "boxplot",
                     "Diagramme en barres" = "bar",
                     "Nuage de points" = "scatter")
        default <- if (n_observations < 30) "boxplot" else if (unique_levels < 20) "bar" else "histogram"
      } else if (var_type %in% c("categorical", "binary")) {
        if (unique_levels == 2) {
          choices <- c("Camembert" = "pie", 
                       "Diagramme en barres" = "bar",
                       "Graphique en beignet" = "donut")
          default <- "pie"
        } else if (unique_levels <= 10) {
          choices <- c("Diagramme en barres" = "bar", 
                       "Camembert" = "pie",
                       "Graphique en beignet" = "donut")
          default <- "bar"
        } else if (unique_levels <= 20) {
          choices <- c("Diagramme en barres" = "bar",
                       "Graphique horizontal" = "bar_horizontal")
          default <- "bar"
        } else {
          choices <- c("Diagramme en barres" = "bar",
                       "Graphique horizontal" = "bar_horizontal")
          default <- "bar_horizontal"
        }
      } else {
        choices <- c("Diagramme en barres" = "bar")
        default <- "bar"
      }
      div(
        style = "background-color: #f8f9fa; padding: 15px; border-radius: 8px; margin: 10px 0;",
        h5("📊 Type de Graphique", style = "color: #495057; margin-bottom: 15px;"),
        radioButtons("plot_type", NULL, 
                     choices = choices, selected = default, inline = FALSE),
        tags$small(
          class = "text-muted",
          icon("info-circle"), " Sélection automatique basée sur le type de variable et les données"
        )
      )
    }, error = function(e) {
      div(class = "alert alert-danger", "❌ Erreur lors de l'analyse de la variable")
    })
  })
  
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
      n_observations <- length(var_data_clean)
      recommendations <- list()
      if (var_type == "continuous") {
        if (n_observations < 30) {
          recommendations$main <- list(
            text = "📊 Peu de données : Boîte à moustaches recommandée",
            color = "#17a2b8",
            icon = "📊"
          )
        } else if (unique_levels < 20) {
          recommendations$main <- list(
            text = "📊 Variable discrète : Diagramme en barres adapté",
            color = "#28a745",
            icon = "📊"
          )
        } else {
          recommendations$main <- list(
            text = "📊 Variable continue : Histogramme optimal",
            color = "#007bff",
            icon = "📊"
          )
        }
        if (n_observations > 1000) {
          recommendations$additional <- list(
            text = "💡 Gros dataset : Considérez un graphique de densité",
            color = "#6c757d",
            icon = "💡"
          )
        }
      } else if (var_type %in% c("categorical", "binary")) {
        if (unique_levels == 2) {
          recommendations$main <- list(
            text = "🥧 Variable binaire : Camembert parfait",
            color = "#28a745",
            icon = "🥧"
          )
        } else if (unique_levels <= 5) {
          recommendations$main <- list(
            text = "📊 Peu de catégories : Graphique en barres ou camembert",
            color = "#ffc107",
            icon = "📊"
          )
        } else if (unique_levels <= 15) {
          recommendations$main <- list(
            text = "📊 Catégories multiples : Diagramme en barres recommandé",
            color = "#fd7e14",
            icon = "📊"
          )
        } else {
          recommendations$main <- list(
            text = "⚠️ Trop de catégories : Graphique horizontal suggéré",
            color = "#dc3545",
            icon = "⚠️"
          )
        }
      }
      if (missing_count > 0) {
        missing_pct <- round(missing_count / length(var_data) * 100, 1)
        if (missing_pct > 20) {
          recommendations$warning <- list(
            text = paste0("⚠️ ", missing_pct, "% de données manquantes - Vérifiez la qualité"),
            color = "#dc3545",
            icon = "⚠️"
          )
        } else {
          recommendations$info <- list(
            text = paste0("ℹ️ ", missing_count, " valeurs manquantes (", missing_pct, "%)"),
            color = "#17a2b8",
            icon = "ℹ️"
          )
        }
      }
      recommendation_cards <- lapply(recommendations, function(rec) {
        div(
          style = paste0("background: linear-gradient(135deg, ", rec$color, "15 0%, ", rec$color, "25 100%); ",
                         "border-left: 4px solid ", rec$color, "; ",
                         "padding: 12px; margin: 8px 0; border-radius: 6px; ",
                         "box-shadow: 0 2px 4px rgba(0,0,0,0.1);"),
          p(rec$text, 
            style = paste0("color: ", rec$color, "; font-weight: 600; margin: 0; font-size: 14px;"))
        )
      })
      if (length(recommendation_cards) > 0) {
        div(
          style = "margin: 15px 0;",
          h5("🎯 Recommandations Intelligentes", style = "color: #495057; margin-bottom: 10px;"),
          do.call(tagList, recommendation_cards)
        )
      }
    }, error = function(e) { 
      div(class = "alert alert-warning", "⚠️ Impossible de générer les recommandations")
    })
  })
  
  generate_enhanced_plot <- function(data, var, plot_type, group_var = NULL, 
                                     title = "", xlab = "", ylab = "", subtitle = "", caption = "",
                                     font_family = "Arial", base_font_size = 12,
                                     title_size = 16, title_style = "bold",
                                     axis_title_size = 14, axis_title_style = "bold",
                                     axis_text_size = 12, axis_text_style = "plain",
                                     legend_text_size = 11, legend_title_style = "bold",
                                     color_mode = "palette", color_palette = "excel",
                                     single_color = "#5B9BD5", custom_colors = NULL,
                                     legend_position = "right", y_axis_type = "counts",
                                     x_axis_rotation = 0, show_grid_major = TRUE,
                                     show_grid_minor = FALSE, background_theme = "white",
                                     margin_top = 20, margin_right = 20, margin_bottom = 20, margin_left = 20) {
    if (is.null(data) || nrow(data) == 0 || !var %in% names(data)) {
      return(NULL)
    }
    var_data <- data[[var]]
    var_data_clean <- var_data[!is.na(var_data)]
    if (length(var_data_clean) == 0) {
      return(NULL)
    }
    # Vérification de la police
    if (!font_family %in% c("Arial", "Times New Roman", "Calibri", "Helvetica", "Georgia", "Verdana")) {
      font_family <- "Arial"
    }
    get_color_palette <- function() {
      if (color_mode == "single") {
        return(single_color)
      } else if (color_mode == "custom" && !is.null(custom_colors) && length(custom_colors) > 0) {
        return(custom_colors)
      } else {
        switch(color_palette,
               "excel" = c("#5B9BD5", "#70AD47", "#FFC000", "#FF6B6B", "#A855F7", 
                           "#06B6D4", "#E74C3C", "#9B59B6", "#F39C12", "#1ABC9C"),
               "viridis" = viridis::viridis(10),
               "set1" = RColorBrewer::brewer.pal(min(9, length(unique(var_data_clean))), "Set1"),
               "set2" = RColorBrewer::brewer.pal(min(8, length(unique(var_data_clean))), "Set2"),
               "dark2" = RColorBrewer::brewer.pal(min(8, length(unique(var_data_clean))), "Dark2"),
               "paired" = RColorBrewer::brewer.pal(min(12, length(unique(var_data_clean))), "Paired"),
               "pastel1" = RColorBrewer::brewer.pal(min(9, length(unique(var_data_clean))), "Pastel1"),
               "accent" = RColorBrewer::brewer.pal(min(8, length(unique(var_data_clean))), "Accent"),
               c("#5B9BD5", "#70AD47", "#FFC000", "#FF6B6B", "#A855F7")
        )
      }
    }
    colors <- get_color_palette()
    base_theme <- theme_minimal(base_size = base_font_size, base_family = font_family) +
      theme(
        plot.title = element_text(size = title_size, face = title_style, hjust = 0.5),
        plot.subtitle = element_text(size = base_font_size + 1, hjust = 0.5),
        plot.caption = element_text(size = base_font_size - 1, hjust = 1),
        axis.title = element_text(size = axis_title_size, face = axis_title_style),
        axis.text = element_text(size = axis_text_size, face = axis_text_style),
        axis.text.x = element_text(angle = x_axis_rotation, hjust = if(x_axis_rotation != 0) 1 else 0.5),
        legend.text = element_text(size = legend_text_size),
        legend.title = element_text(size = legend_text_size, face = legend_title_style),
        legend.position = if(legend_position == "none") "none" else legend_position,
        plot.margin = margin(t = margin_top, r = margin_right, b = margin_bottom, l = margin_left),
        panel.grid.major = if(show_grid_major) element_line(color = "gray90", size = 0.5) else element_blank(),
        panel.grid.minor = if(show_grid_minor) element_line(color = "gray95", size = 0.25) else element_blank()
      )
    if (background_theme == "dark") {
      base_theme <- base_theme + theme(
        plot.background = element_rect(fill = "#2E3440", color = NA),
        panel.background = element_rect(fill = "#3B4252", color = NA),
        text = element_text(color = "white"),
        axis.text = element_text(color = "white"),
        legend.text = element_text(color = "white"),
        legend.title = element_text(color = "white")
      )
    } else if (background_theme == "light_gray") {
      base_theme <- base_theme + theme(
        plot.background = element_rect(fill = "#F8F9FA", color = NA),
        panel.background = element_rect(fill = "#FFFFFF", color = NA)
      )
    }
    p <- tryCatch({
      switch(plot_type,
             "histogram" = {
               ggplot(data, aes(x = .data[[var]])) +
                 geom_histogram(fill = colors[1], color = "white", alpha = 0.8, bins = 30) +
                 labs(title = title, subtitle = subtitle, caption = caption,
                      x = if(xlab == "") var else xlab,
                      y = if(ylab == "") "Effectifs" else ylab) +
                 base_theme
             },
             "density" = {
               ggplot(data, aes(x = .data[[var]])) +
                 geom_density(fill = colors[1], color = colors[1], alpha = 0.6, size = 1.2) +
                 labs(title = title, subtitle = subtitle, caption = caption,
                      x = if(xlab == "") var else xlab,
                      y = if(ylab == "") "Densité" else ylab) +
                 base_theme
             },
             "boxplot" = {
               if (!is.null(group_var) && group_var != "" && group_var %in% names(data)) {
                 ggplot(data, aes(x = .data[[group_var]], y = .data[[var]], fill = .data[[group_var]])) +
                   geom_boxplot(alpha = 0.8, outlier.color = colors[1]) +
                   scale_fill_manual(values = colors) +
                   labs(title = title, subtitle = subtitle, caption = caption,
                        x = if(xlab == "") group_var else xlab,
                        y = if(ylab == "") var else ylab) +
                   base_theme
               } else {
                 ggplot(data, aes(y = .data[[var]])) +
                   geom_boxplot(fill = colors[1], alpha = 0.8, outlier.color = colors[1]) +
                   labs(title = title, subtitle = subtitle, caption = caption,
                        x = if(xlab == "") "" else xlab,
                        y = if(ylab == "") var else ylab) +
                   base_theme
               }
             },
             "bar" = {
               if (detect_var_type(var_data_clean) == "continuous") {
                 data_plot <- data %>%
                   filter(!is.na(.data[[var]])) %>%
                   mutate(var_cut = cut(.data[[var]], breaks = 10, include.lowest = TRUE))
                 if (y_axis_type == "percent") {
                   data_plot <- data_plot %>%
                     count(var_cut) %>%
                     mutate(percent = n / sum(n) * 100)
                   ggplot(data_plot, aes(x = var_cut, y = percent)) +
                     geom_col(fill = colors[1], alpha = 0.8) +
                     labs(title = title, subtitle = subtitle, caption = caption,
                          x = if(xlab == "") var else xlab,
                          y = if(ylab == "") "Pourcentage (%)" else ylab) +
                     base_theme
                 } else {
                   ggplot(data_plot, aes(x = var_cut)) +
                     geom_bar(fill = colors[1], alpha = 0.8) +
                     labs(title = title, subtitle = subtitle, caption = caption,
                          x = if(xlab == "") var else xlab,
                          y = if(ylab == "") "Effectifs" else ylab) +
                     base_theme
                 }
               } else {
                 if (!is.null(group_var) && group_var != "" && group_var %in% names(data)) {
                   if (y_axis_type == "percent") {
                     data_plot <- data %>%
                       filter(!is.na(.data[[var]])) %>%
                       count(.data[[var]], .data[[group_var]]) %>%
                       group_by(.data[[group_var]]) %>%
                       mutate(percent = n / sum(n) * 100)
                     ggplot(data_plot, aes(x = .data[[var]], y = percent, fill = .data[[group_var]])) +
                       geom_col(position = "dodge", alpha = 0.8) +
                       scale_fill_manual(values = colors) +
                       labs(title = title, subtitle = subtitle, caption = caption,
                            x = if(xlab == "") var else xlab,
                            y = if(ylab == "") "Pourcentage (%)" else ylab,
                            fill = group_var) +
                       base_theme
                   } else {
                     ggplot(data, aes(x = .data[[var]], fill = .data[[group_var]])) +
                       geom_bar(position = "dodge", alpha = 0.8) +
                       scale_fill_manual(values = colors) +
                       labs(title = title, subtitle = subtitle, caption = caption,
                            x = if(xlab == "") var else xlab,
                            y = if(ylab == "") "Effectifs" else ylab,
                            fill = group_var) +
                       base_theme
                   }
                 } else {
                   if (y_axis_type == "percent") {
                     data_plot <- data %>%
                       filter(!is.na(.data[[var]])) %>%
                       count(.data[[var]]) %>%
                       mutate(percent = n / sum(n) * 100)
                     ggplot(data_plot, aes(x = .data[[var]], y = percent)) +
                       geom_col(fill = colors[1], alpha = 0.8) +
                       labs(title = title, subtitle = subtitle, caption = caption,
                            x = if(xlab == "") var else xlab,
                            y = if(ylab == "") "Pourcentage (%)" else ylab) +
                       base_theme
                   } else {
                     ggplot(data, aes(x = .data[[var]])) +
                       geom_bar(fill = colors[1], alpha = 0.8) +
                       labs(title = title, subtitle = subtitle, caption = caption,
                            x = if(xlab == "") var else xlab,
                            y = if(ylab == "") "Effectifs" else ylab) +
                       base_theme
                   }
                 }
               }
             },
             "bar_horizontal" = {
               data_plot <- data %>%
                 filter(!is.na(.data[[var]])) %>%
                 count(.data[[var]], sort = TRUE)
               if (y_axis_type == "percent") {
                 data_plot <- data_plot %>%
                   mutate(percent = n / sum(n) * 100)
                 ggplot(data_plot, aes(x = percent, y = reorder(.data[[var]], percent))) +
                   geom_col(fill = colors[1], alpha = 0.8) +
                   labs(title = title, subtitle = subtitle, caption = caption,
                        x = if(ylab == "") "Pourcentage (%)" else ylab,
                        y = if(xlab == "") var else xlab) +
                   base_theme
               } else {
                 ggplot(data_plot, aes(x = n, y = reorder(.data[[var]], n))) +
                   geom_col(fill = colors[1], alpha = 0.8) +
                   labs(title = title, subtitle = subtitle, caption = caption,
                        x = if(ylab == "") "Effectifs" else ylab,
                        y = if(xlab == "") var else xlab) +
                   base_theme
               }
             },
             "pie" = {
               data_plot <- data %>%
                 filter(!is.na(.data[[var]])) %>%
                 count(.data[[var]]) %>%
                 mutate(percent = n / sum(n) * 100,
                        label = paste0(.data[[var]], "\n", round(percent, 1), "%"))
               ggplot(data_plot, aes(x = "", y = n, fill = .data[[var]])) +
                 geom_col(width = 1, alpha = 0.8) +
                 coord_polar("y", start = 0) +
                 scale_fill_manual(values = colors) +
                 labs(title = title, subtitle = subtitle, caption = caption, fill = var) +
                 base_theme +
                 theme(axis.text = element_blank(),
                       axis.title = element_blank(),
                       panel.grid = element_blank())
             },
             "donut" = {
               data_plot <- data %>%
                 filter(!is.na(.data[[var]])) %>%
                 count(.data[[var]]) %>%
                 mutate(percent = n / sum(n) * 100,
                        label = paste0(.data[[var]], "\n", round(percent, 1), "%"))
               ggplot(data_plot, aes(x = 2, y = n, fill = .data[[var]])) +
                 geom_col(width = 1, alpha = 0.8) +
                 coord_polar("y", start = 0) +
                 xlim(0.5, 2.5) +
                 scale_fill_manual(values = colors) +
                 labs(title = title, subtitle = subtitle, caption = caption, fill = var) +
                 base_theme +
                 theme(axis.text = element_blank(),
                       axis.title = element_blank(),
                       panel.grid = element_blank())
             },
             {
               ggplot(data, aes(x = .data[[var]])) +
                 geom_bar(fill = colors[1], alpha = 0.8) +
                 labs(title = title, subtitle = subtitle, caption = caption,
                      x = if(xlab == "") var else xlab,
                      y = if(ylab == "") "Effectifs" else ylab) +
                 base_theme
             }
      )
    }, error = function(e) {
      showNotification(paste("Erreur lors de la génération du graphique :", e$message), type = "error")
      NULL
    })
    return(p)
  }
  
  current_plot <- reactive({
    req(input$generate_plot, current_data(), input$plot_var, input$plot_type)
    custom_colors <- NULL
    if (input$color_mode == "custom" && !is.null(input$n_colors)) {
      custom_colors <- sapply(1:input$n_colors, function(i) {
        input[[paste0("custom_color_", i)]] %||% "#5B9BD5"
      })
    }
    plot_obj <- generate_enhanced_plot(
      data = current_data(),
      var = input$plot_var,
      plot_type = input$plot_type,
      group_var = if(input$group_var == "") NULL else input$group_var,
      title = input$plot_title,
      xlab = input$plot_xlab,
      ylab = input$plot_ylab,
      subtitle = input$plot_subtitle,
      caption = input$plot_caption,
      font_family = input$font_family,
      base_font_size = input$base_font_size,
      title_size = input$title_size,
      title_style = input$title_style,
      axis_title_size = input$axis_title_size,
      axis_title_style = input$axis_title_style,
      axis_text_size = input$axis_text_size,
      axis_text_style = input$axis_text_style,
      legend_text_size = input$legend_text_size,
      legend_title_style = input$legend_title_style,
      color_mode = input$color_mode,
      color_palette = input$color_palette,
      single_color = input$single_color,
      custom_colors = custom_colors,
      legend_position = input$legend_position,
      y_axis_type = input$y_axis_type,
      x_axis_rotation = input$x_axis_rotation,
      show_grid_major = input$show_grid_major,
      show_grid_minor = input$show_grid_minor,
      background_theme = input$background_theme,
      margin_top = input$margin_top,
      margin_right = input$margin_right,
      margin_bottom = input$margin_bottom,
      margin_left = input$margin_left
    )
    if (input$use_plotly && !is.null(plot_obj)) {
      tryCatch({
        plot_obj <- plotly::ggplotly(plot_obj) %>%
          plotly::config(displayModeBar = TRUE, displaylogo = FALSE)
      }, error = function(e) {
        showNotification("Erreur lors de la conversion en Plotly", type = "error")
        plot_obj <- NULL
      })
    }
    return(plot_obj)
  })
  
  output$current_plot_static <- renderPlot({
    req(current_plot())
    if (!inherits(current_plot(), "plotly")) {
      current_plot()
    }
  })
  
  output$current_plot_plotly <- renderPlotly({
    req(current_plot())
    if (inherits(current_plot(), "plotly")) {
      current_plot()
    }
  })
  
  output$plot_detailed_info <- renderUI({
    req(current_plot(), current_data(), input$plot_var)
    data <- current_data()
    var_data <- data[[input$plot_var]]
    var_data_clean <- var_data[!is.na(var_data)]
    stats_info <- if (detect_var_type(var_data_clean) == "continuous") {
      paste0("📊 Statistiques : Min=", round(min(var_data_clean), 2), 
             " | Max=", round(max(var_data_clean), 2),
             " | Moyenne=", round(mean(var_data_clean), 2),
             " | Médiane=", round(median(var_data_clean), 2))
    } else {
      n_levels <- length(unique(var_data_clean))
      paste0("📊 Statistiques : ", n_levels, " catégories | Mode: ", 
             names(sort(table(var_data_clean), decreasing = TRUE))[1])
    }
    data_info <- paste0("📈 Données : ", length(var_data_clean), " observations | ",
                        sum(is.na(var_data)), " valeurs manquantes")
    plot_params <- paste0("⚙️ Paramètres : Type=", input$plot_type, 
                          " | Couleurs=", input$color_mode,
                          " | Police=", input$font_family)
    div(
      style = "background-color: #f8f9fa; padding: 15px; border-radius: 8px; margin: 10px 0;",
      h5("📋 Informations du Graphique", style = "color: #495057; margin-bottom: 10px;"),
      tags$p(stats_info, style = "margin: 5px 0; color: #6c757d;"),
      tags$p(data_info, style = "margin: 5px 0; color: #6c757d;"),
      tags$p(plot_params, style = "margin: 5px 0; color: #6c757d;")
    )
  })
  
  observeEvent(input$add_to_export, {
    req(current_plot())
    new_plot <- list(
      plot = current_plot(),
      variable = input$plot_var,
      type = input$plot_type,
      title = input$plot_title,
      timestamp = Sys.time(),
      settings = list(
        font_family = input$font_family,
        color_mode = input$color_mode,
        color_palette = input$color_palette
      )
    )
    current_history <- plot_history()
    current_history[[length(current_history) + 1]] <- new_plot
    plot_history(current_history)
    showNotification("Graphique ajouté à l'historique", type = "message")
  })
  
  output$enhanced_visualization_list <- renderUI({
    history <- plot_history()
    if (length(history) == 0) {
      return(div(class = "alert alert-info", "Aucun graphique dans l'historique"))
    }
    filtered_history <- history
    if (input$filter_plot_type != "all") {
      filtered_history <- filtered_history[sapply(filtered_history, function(x) x$type == input$filter_plot_type)]
    }
    if (input$filter_variable != "all") {
      filtered_history <- filtered_history[sapply(filtered_history, function(x) x$variable == input$filter_variable)]
    }
    history_cards <- lapply(1:length(filtered_history), function(i) {
      plot_item <- filtered_history[[i]]
      div(
        class = "col-md-4",
        style = "margin-bottom: 20px;",
        div(
          class = "card",
          style = "border: 1px solid #dee2e6; border-radius: 8px;",
          div(
            class = "card-header",
            style = "background-color: #f8f9fa; border-bottom: 1px solid #dee2e6;",
            h6(plot_item$title, style = "margin: 0; color: #495057;")
          ),
          div(
            class = "card-body",
            style = "padding: 15px;",
            tags$p(paste("Variable:", plot_item$variable), style = "margin: 5px 0; color: #6c757d;"),
            tags$p(paste("Type:", plot_item$type), style = "margin: 5px 0; color: #6c757d;"),
            tags$p(paste("Créé:", format(plot_item$timestamp, "%H:%M")), style = "margin: 5px 0; color: #6c757d;"),
            div(
              style = "margin-top: 10px;",
              actionButton(paste0("view_plot_", i), "👁️ Voir", class = "btn-sm btn-info", style = "margin-right: 5px;"),
              actionButton(paste0("delete_plot_", i), "🗑️ Supprimer", class = "btn-sm btn-danger")
            )
          )
        )
      )
    })
    div(class = "row", do.call(tagList, history_cards))
  })
  
  observeEvent(input$clear_history, {
    plot_history(list())
    showNotification("Historique vidé", type = "warning")
  })
  
  observeEvent(input$reset_settings, {
    updateTextInput(session, "plot_title", value = "Graphique univarié")
    updateTextInput(session, "plot_xlab", value = "")
    updateTextInput(session, "plot_ylab", value = "")
    updateTextInput(session, "plot_subtitle", value = "")
    updateTextInput(session, "plot_caption", value = "")
    updateSelectInput(session, "font_family", selected = "Arial")
    updateNumericInput(session, "base_font_size", value = 12)
    updateSelectInput(session, "color_mode", selected = "palette")
    updateSelectInput(session, "color_palette", selected = "excel")
    updateSelectInput(session, "legend_position", selected = "right")
    updateCheckboxInput(session, "use_plotly", value = FALSE)
    showNotification("Paramètres réinitialisés", type = "info")
  })
  
  generate_r_code <- function() {
    req(current_data(), input$plot_var)
    code <- paste0(
      "# Code R généré automatiquement\n",
      "library(ggplot2)\n",
      "library(dplyr)\n\n",
      "# Génération du graphique\n",
      "p <- ggplot(data, aes(x = ", input$plot_var, ")) +\n"
    )
    if (input$plot_type == "histogram") {
      code <- paste0(code, "  geom_histogram(fill = '", input$single_color, "', alpha = 0.8) +\n")
    } else if (input$plot_type == "bar") {
      code <- paste0(code, "  geom_bar(fill = '", input$single_color, "', alpha = 0.8) +\n")
    } else if (input$plot_type == "density") {
      code <- paste0(code, "  geom_density(fill = '", input$single_color, "', alpha = 0.6) +\n")
    } else if (input$plot_type == "boxplot") {
      code <- paste0(code, "  geom_boxplot(fill = '", input$single_color, "', alpha = 0.8) +\n")
    } else if (input$plot_type == "pie") {
      code <- paste0(code, "  geom_bar(fill = '", input$single_color, "', alpha = 0.8) +\n  coord_polar('y', start = 0) +\n")
    } else if (input$plot_type == "donut") {
      code <- paste0(code, "  geom_bar(fill = '", input$single_color, "', alpha = 0.8) +\n  coord_polar('y', start = 0) +\n  xlim(0.5, 2.5) +\n")
    } else if (input$plot_type == "bar_horizontal") {
      code <- paste0(code, "  geom_bar(fill = '", input$single_color, "', alpha = 0.8) +\n  coord_flip() +\n")
    }
    code <- paste0(code, 
                   "  labs(\n",
                   "    title = '", input$plot_title, "',\n",
                   "    x = '", if(is.null(input$plot_xlab) || input$plot_xlab == "") input$plot_var else input$plot_xlab, "',\n",
                   "    y = '", if(is.null(input$plot_ylab) || input$plot_ylab == "") "Effectifs" else input$plot_ylab, "'\n",
                   "  ) +\n",
                   "  theme_minimal(base_size = ", input$base_font_size, ") +\n",
                   "  theme(\n",
                   "    plot.title = element_text(size = ", input$title_size, ", face = '", input$title_style, "'),\n",
                   "    legend.position = '", input$legend_position, "'\n",
                   "  )\n\n",
                   "# Affichage du graphique\n",
                   "print(p)\n"
    )
    return(code)
  }
  
  observeEvent(input$copy_code, {
    code <- generate_r_code()
    showModal(modalDialog(
      title = "Code R généré",
      tags$pre(code),
      easyClose = TRUE,
      footer = tagList(
        modalButton("Fermer"),
        actionButton("copy_to_clipboard", "Copier", class = "btn-primary")
      )
    ))
  })
  
  output$download_plot <- downloadHandler(
    filename = function() {
      paste0(input$plot_filename, ".", input$plot_format)
    },
    content = function(file) {
      req(current_plot())
      tryCatch({
        if (input$plot_format == "png") {
          ggsave(file, plot = current_plot(), width = input$export_width, height = input$export_height, 
                 dpi = input$export_dpi, device = "png")
        } else if (input$plot_format == "jpg") {
          ggsave(file, plot = current_plot(), width = input$export_width, height = input$export_height, 
                 dpi = input$export_dpi, device = "jpeg")
        } else if (input$plot_format == "pdf") {
          ggsave(file, plot = current_plot(), width = input$export_width, height = input$export_height, 
                 device = "pdf")
        } else if (input$plot_format == "svg") {
          ggsave(file, plot = current_plot(), width = input$export_width, height = input$export_height, 
                 device = "svg")
        } else if (input$plot_format == "eps") {
          ggsave(file, plot = current_plot(), width = input$export_width, height = input$export_height, 
                 device = "eps")
        }
      }, error = function(e) {
        showNotification(paste("Erreur lors de l'exportation :", e$message), type = "error")
      })
    }
  )
  
  output$export_preview_list <- renderUI({
    plots <- plot_history()
    if (length(plots) == 0) {
      return(div(class = "alert alert-info", "Aucun graphique généré pour le moment."))
    }
    preview_items <- lapply(seq_along(plots), function(i) {
      item <- plots[[i]]
      div(
        class = "preview-item",
        style = "border: 1px solid #ddd; padding: 15px; margin-bottom: 10px; border-radius: 5px; background-color: #f9f9f9;",
        fluidRow(
          column(8,
                 h5(paste("Graphique", i)),
                 p(strong("Titre: "), item$title),
                 p(strong("Type: "), item$type),
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
      return(div(class = "alert alert-info", "Sélectionnez un graphique pour une prévisualisation."))
    }
    if (inherits(selected$plot, "plotly")) {
      plotlyOutput("preview_plotly", height = "400px")
    } else {
      plotOutput("preview_static", height = "400px")
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
  
  output$download_html <- downloadHandler(
    filename = function() { paste0("rapport_", Sys.Date(), ".html") },
    content = function(file) {
      plots <- plot_history()
      selected_items <- sapply(seq_along(plots), function(i) {
        if (is.null(input[[paste0("export_item_", i)]])) FALSE else input[[paste0("export_item_", i)]]
      })
      if (!any(selected_items)) {
        showNotification("Aucun graphique sélectionné pour l'exportation.", type = "warning")
        return()
      }
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
      for (i in seq_along(plots)) {
        if (selected_items[i]) {
          plot_info <- plots[[i]]
          temp_file <- tempfile(fileext = ".png")
          tryCatch({
            if (inherits(plot_info$plot, "plotly")) {
              plotly::export(plot_info$plot, file = temp_file)
            } else {
              ggsave(temp_file, plot = plot_info$plot, device = "png", width = 12, height = 8, dpi = 300)
            }
            img_data <- base64enc::base64encode(readBin(temp_file, "raw", file.info(temp_file)$size))
            html_content <- c(html_content,
                              "<h2>", paste("Graphique ", i, ": ", plot_info$title), "</h2>",
                              "<img src='data:image/png;base64,", img_data, "' alt='", plot_info$title, "'>")
            unlink(temp_file)
          }, error = function(e) {
            showNotification(paste("Erreur lors de l'exportation du graphique", i, ":", e$message), type = "error")
          })
        }
      }
      html_content <- c(html_content, "</body></html>")
      writeLines(html_content, file)
      showNotification("Rapport HTML généré avec succès !", type = "message")
    }
  )
  
  output$download_csv <- downloadHandler(
    filename = function() { paste0("données_", Sys.Date(), ".csv") },
    content = function(file) {
      req(current_data())
      tryCatch({
        write_csv(current_data(), file)
        showNotification("Fichier CSV téléchargé avec succès !", type = "message")
      }, error = function(e) {
        showNotification(paste("Erreur lors de l'exportation CSV :", e$message), type = "error")
      })
    }
  )
  
  observe({
    history <- plot_history()
    if (length(history) > 0) {
      plot_types <- unique(sapply(history, function(x) x$type))
      variables <- unique(sapply(history, function(x) x$variable))
      updateSelectInput(session, "filter_plot_type", 
                        choices = c("Tous" = "all", setNames(plot_types, plot_types)))
      updateSelectInput(session, "filter_variable", 
                        choices = c("Toutes" = "all", setNames(variables, variables)))
    }
  })
  
  observe({
    history <- plot_history()
    lapply(seq_along(history), function(i) {
      observeEvent(input[[paste0("view_plot_", i)]], {
        selected_preview(history[[i]])
      })
      observeEvent(input[[paste0("delete_plot_", i)]], {
        current_history <- plot_history()
        current_history <- current_history[-i]
        plot_history(current_history)
        showNotification("Graphique supprimé de l'historique", type = "warning")
      })
    })
  })
  
  observeEvent(input$preview_settings, {
    showModal(modalDialog(
      title = "Aperçu des Paramètres",
      tags$div(
        style = "font-size: 14px;",
        tags$p(strong("Variable principale: "), input$plot_var),
        tags$p(strong("Type de graphique: "), input$plot_type),
        tags$p(strong("Titre: "), input$plot_title),
        tags$p(strong("Police: "), input$font_family),
        tags$p(strong("Mode de couleur: "), input$color_mode),
        if (input$color_mode == "palette") {
          tags$p(strong("Palette: "), input$color_palette)
        } else if (input$color_mode == "single") {
          tags$p(strong("Couleur: "), input$single_color)
        } else {
          tags$p(strong("Nombre de couleurs personnalisées: "), input$n_colors)
        }
      ),
      easyClose = TRUE,
      footer = modalButton("Fermer")
    ))
  })
}

# Lancement de l'application
shinyApp(ui = ui, server = server)
