# gtsummary Studio Pro - Version Améliorée

# Charger les bibliothèques nécessaires
required_packages <- c("shiny", "shinydashboard", "DT", "gtsummary", "gt", "dplyr", "readr", "readxl", "shinyWidgets", "shinycssloaders", "flextable", "officer")
missing_packages <- required_packages[!sapply(required_packages, requireNamespace, quietly = TRUE)]
if (length(missing_packages) > 0) {
  stop(paste("Les packages suivants doivent être installés :", paste(missing_packages, collapse = ", ")))
}

# Charger les bibliothèques
lapply(required_packages, library, character.only = TRUE)

# Configurer le thème gtsummary en français
theme_gtsummary_language(language = "fr", decimal.mark = ",", big.mark = " ", set_theme = TRUE)

# Fonctions utilitaires
detect_var_type <- function(x) {
  if (is.numeric(x)) {
    unique_vals <- length(unique(x[!is.na(x)]))
    if (unique_vals <= 10 && all(x[!is.na(x)] %% 1 == 0, na.rm = TRUE)) {
      return("categorical")
    }
    return("continuous")
  } else if (is.factor(x) || is.character(x)) {
    return("categorical")
  } else if (is.logical(x)) {
    return("binary")
  }
  return("unknown")
}

get_available_dataframes <- function() {
  all_objects <- ls(envir = .GlobalEnv)
  dataframes <- character(0)
  for (obj_name in all_objects) {
    obj <- get(obj_name, envir = .GlobalEnv)
    if (is.data.frame(obj) && nrow(obj) > 0) {
      dataframes <- c(dataframes, obj_name)
    }
  }
  return(dataframes)
}

# Interface utilisateur
ui <- dashboardPage(
  dashboardHeader(
    title = tags$div(
      class = "header-title",
      tags$img(src = "https://via.placeholder.com/24", style = "margin-right: 8px; vertical-align: middle;"),
      "gtsummary Studio Pro"
    ),
    titleWidth = 280
  ),

  dashboardSidebar(
    width = 280,
    tags$style(HTML("
      /* Global Styles */
      .content-wrapper {
        background-color: #F7F9FB;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
        color: #1A202C;
      }
      .main-header {
        background: linear-gradient(90deg, #2B6CB0 0%, #2C5282 100%);
        border-bottom: 1px solid rgba(255,255,255,0.1);
      }
      .main-header .logo {
        font-weight: 600;
        font-size: 18px;
        padding: 0 15px;
        transition: all 0.3s ease;
      }
      .header-title img {
        filter: brightness(0) invert(1);
      }
      .sidebar {
        background-color: #1A202C;
        box-shadow: 2px 0 8px rgba(0,0,0,0.1);
      }
      .sidebar-menu .treeview-menu {
        background-color: #2D3748;
      }
      .sidebar-menu li a {
        color: #E2E8F0;
        font-weight: 500;
        padding: 12px 15px;
        transition: all 0.3s ease;
      }
      .sidebar-menu li a:hover {
        background-color: #2B6CB0;
        color: #FFFFFF;
        transform: translateX(5px);
      }
      .sidebar-menu li.active a {
        background-color: #2B6CB0;
        color: #FFFFFF;
        border-left: 4px solid #68D391;
      }
      .sidebar-menu .menu-icon {
        margin-right: 10px;
        font-size: 16px;
      }
      .box {
        border-radius: 8px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.08);
        border: none;
        background-color: #FFFFFF;
        transition: transform 0.2s ease, box-shadow 0.2s ease;
      }
      .box:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 16px rgba(0,0,0,0.12);
      }
      .box-header {
        background: linear-gradient(90deg, #EDF2F7 0%, #E2E8F0 100%);
        border-radius: 8px 8px 0 0;
        padding: 12px 16px;
        border-bottom: none;
      }
      .box-title {
        font-size: 16px;
        font-weight: 600;
        color: #2D3748;
      }
      .btn {
        border-radius: 6px;
        padding: 8px 16px;
        font-weight: 500;
        transition: all 0.3s ease;
      }
      .btn-primary {
        background-color: #2B6CB0;
        border-color: #2B6CB0;
      }
      .btn-primary:hover {
        background-color: #2C5282;
        border-color: #2C5282;
        transform: translateY(-1px);
      }
      .btn-success {
        background-color: #68D391;
        border-color: #68D391;
      }
      .btn-success:hover {
        background-color: #48BB78;
        border-color: #48BB78;
        transform: translateY(-1px);
      }
      .btn-info {
        background-color: #4FD1C5;
        border-color: #4FD1C5;
      }
      .btn-info:hover {
        background-color: #38B2AC;
        border-color: #38B2AC;
        transform: translateY(-1px);
      }
      .btn-warning {
        background-color: #ECC94B;
        border-color: #ECC94B;
      }
      .btn-warning:hover {
        background-color: #D69E2E;
        border-color: #D69E2E;
        transform: translateY(-1px);
      }
      .form-control {
        border-radius: 6px;
        border: 1px solid #E2E8F0;
        padding: 8px 12px;
        transition: all 0.3s ease;
      }
      .form-control:focus {
        border-color: #2B6CB0;
        box-shadow: 0 0 0 3px rgba(43,108,176,0.2);
      }
      .shiny-input-container {
        margin-bottom: 16px;
      }
      .control-label {
        font-size: 14px;
        font-weight: 500;
        color: #4A5568;
        margin-bottom: 6px;
      }
      .alert-info {
        background-color: #EBF8FF;
        border: 1px solid #BEE3F8;
        color: #2B6CB0;
        border-radius: 6px;
        padding: 12px;
        font-size: 14px;
      }
      .error-message {
        color: #E53E3E;
        font-weight: 500;
        font-size: 14px;
      }
      .success-message {
        color: #48BB78;
        font-weight: 500;
        font-size: 14px;
      }
      .dataTables_wrapper .dataTable {
        border-radius: 6px;
        border: none;
        background-color: #FFFFFF;
      }
      .dataTables_wrapper .dataTable th {
        background-color: #EDF2F7;
        color: #2D3748;
        font-weight: 600;
        padding: 12px;
      }
      .dataTables_wrapper .dataTable td {
        padding: 10px;
        border-bottom: 1px solid #E2E8F0;
      }
      .dataTables_wrapper .dataTable tr:hover {
        background-color: #F7FAFC;
      }
      h1, h2, h3, h4, h5 {
        color: #2D3748;
        font-weight: 600;
      }
      h4 { font-size: 18px; }
      h5 { font-size: 16px; }
      p { font-size: 14px; line-height: 1.6; color: #4A5568; }
      .reference-section {
        background-color: #F7FAFC;
        padding: 12px;
        border-radius: 6px;
        margin-bottom: 12px;
        border: 1px solid #E2E8F0;
      }
      .shiny-spinner-output-container .spinner {
        opacity: 0.7;
      }
      @media (max-width: 768px) {
        .sidebar {
          width: 200px !important;
        }
        .main-header .logo {
          font-size: 16px;
        }
        .box {
          margin-bottom: 16px;
        }
      }
    ")),

    sidebarMenu(
      menuItem("📊 Données", tabName = "data", icon = tags$i(class = "fas fa-database menu-icon")),
      menuItem("🎯 Modalités", tabName = "reference", icon = tags$i(class = "fas fa-cog menu-icon")),
      menuItem("📈 Descriptif", tabName = "descriptive", icon = tags$i(class = "fas fa-table menu-icon")),
      menuItem("🔍 Croisé", tabName = "crosstab", icon = tags$i(class = "fas fa-th menu-icon")),
      menuItem("📥 Export", tabName = "export", icon = tags$i(class = "fas fa-download menu-icon"))
    )
  ),

  dashboardBody(
    tabItems(
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
                                 class = "btn-primary", style = "margin-top: 16px;")
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
      tabItem(tabName = "reference",
              fluidRow(
                box(title = "🎯 Gestion des Modalités de Référence", status = "primary", solidHeader = TRUE, width = 12,
                    div(class = "alert-info",
                        "💡 Définissez les modalités de référence pour vos variables catégorielles. Ces paramètres seront utilisés dans toutes les analyses."),
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
                                     p("Veuillez d'abord charger des données dans l'onglet 'Données'.",
                                       class = "error-message"))
                )
              )
      ),
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
      tabItem(tabName = "crosstab",
              fluidRow(
                column(4,
                       box(title = "⚙️ Configuration", status = "primary", solidHeader = TRUE, width = 12,
                           selectInput("cross_row", "Variable en ligne:", choices = NULL),
                           selectInput("cross_col", "Variable en colonne:", choices = NULL),
                           checkboxInput("cross_percent", "Afficher les pourcentages", TRUE),
                           radioButtons("percent_type", "Type de pourcentage:",
                                        choices = list("Ligne" = "row", "Colonne" = "col", "Total" = "cell"),
                                        selected = "row"),
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
      tabItem(tabName = "export",
              fluidRow(
                box(title = "📥 Export des Tableaux", status = "primary", solidHeader = TRUE, width = 12,
                    div(class = "alert-info",
                        "💡 Tous les tableaux générés durant votre session seront inclus dans le rapport Word."),
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
                    ),
                    br(),
                    h4("📋 Tableaux inclus dans le rapport:"),
                    htmlOutput("tables_summary")
                )
              )
      )
    )
  )
)

# Serveur
server <- function(input, output, session) {

  # Variables réactives
  current_data <- reactiveVal(NULL)
  original_data <- reactiveVal(NULL)
  current_table <- reactiveVal(NULL)
  all_tables <- reactiveVal(list())
  reference_levels <- reactiveVal(list())

  # Indicateur de données chargées
  output$data_loaded <- reactive({
    !is.null(current_data())
  })
  outputOptions(output, "data_loaded", suspendWhenHidden = FALSE)

  # Mise à jour des choix de data.frames
  update_dataframe_choices <- function() {
    available_dfs <- get_available_dataframes()
    if (length(available_dfs) > 0) {
      choices <- setNames(available_dfs,
                          paste0(available_dfs, " (",
                                 sapply(available_dfs, function(x) {
                                   df <- get(x, envir = .GlobalEnv)
                                   paste0(nrow(df), " obs, ", ncol(df), " vars")
                                 }), ")"))
      updateSelectInput(session, "r_dataframe", choices = choices)
    } else {
      updateSelectInput(session, "r_dataframe",
                        choices = list("Aucun data.frame disponible" = ""))
    }
  }

  observe({
    if (input$data_source == "r_env") {
      update_dataframe_choices()
    }
  })

  observeEvent(input$refresh_dataframes, {
    update_dataframe_choices()
    showNotification("Liste des data.frames actualisée !", type = "message")
  })

  # Chargement des données
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
          showNotification("Le data.frame sélectionné n'existe plus dans l'environnement.", type = "error")
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

  # Génération des contrôles pour modalités de référence
  output$reference_controls <- renderUI({
    req(current_data())
    data <- current_data()
    cat_vars <- names(data)[sapply(data, function(x) is.factor(x) || is.character(x) || is.logical(x))]
    if (length(cat_vars) == 0) {
      return(p("Aucune variable catégorielle détectée dans vos données.", class = "error-message"))
    }
    controls <- lapply(cat_vars, function(var) {
      levels <- unique(data[[var]][!is.na(data[[var]])])
      if (length(levels) > 1) {
        current_ref <- reference_levels()[[var]]
        if (is.null(current_ref) || !current_ref %in% levels) {
          current_ref <- levels[1]
        }
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

  # Application des modalités de référence
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
          if (is.character(data[[var]])) {
            data[[var]] <- factor(data[[var]])
          }
          if (is.factor(data[[var]])) {
            data[[var]] <- relevel(data[[var]], ref = input[[ref_input_id]])
          }
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

  # Mise à jour des choix de variables
  observe({
    req(current_data())
    data <- current_data()
    all_vars <- names(data)
    cat_vars <- all_vars[sapply(data, function(x) is.factor(x) || is.character(x) || is.logical(x))]
    updateSelectInput(session, "desc_variables",
                      choices = all_vars,
                      selected = all_vars[1:min(5, length(all_vars))])
    updateSelectInput(session, "desc_by",
                      choices = c("Aucun" = "", setNames(cat_vars, cat_vars)))
    updateSelectInput(session, "cross_row", choices = cat_vars)
    updateSelectInput(session, "cross_col", choices = cat_vars)
  })

  # Informations sur les données
  output$data_info <- renderUI({
    req(current_data())
    data <- current_data()
    n_obs <- nrow(data)
    n_vars <- ncol(data)
    n_numeric <- sum(sapply(data, is.numeric))
    n_categorical <- sum(sapply(data, function(x) is.factor(x) || is.character(x)))
    n_missing <- sum(is.na(data))
    HTML(paste(
      "<h5>📈 Résumé des données</h5>",
      "<ul>",
      paste0("<li><strong>Observations:</strong> ", n_obs, "</li>"),
      paste0("<li><strong>Variables:</strong> ", n_vars, "</li>"),
      paste0("<li><strong>Numériques:</strong> ", n_numeric, "</li>"),
      paste0("<li><strong>Catégorielles:</strong> ", n_categorical, "</li>"),
      paste0("<li><strong>Valeurs manquantes:</strong> ", n_missing, "</li>"),
      "</ul>"
    ))
  })

  # Aperçu des données
  output$data_preview <- DT::renderDataTable({
    req(current_data())
    datatable(current_data(),
              options = list(scrollX = TRUE, pageLength = 10, dom = 'tip'))
  })

  # Fonction pour ajouter un tableau à la liste
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

  # Génération du tableau descriptif
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
        if (input$desc_overall) {
          tbl <- tbl %>% add_overall(last = TRUE)
        }
        if (input$add_p_values) {
          tbl <- tbl %>%
            add_p() %>%
            bold_p()
        }
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

  # Génération du tableau croisé
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
          percent = if(input$cross_percent) input$percent_type else "none"
        ) %>%
        bold_labels()
      if (input$add_chi2) {
        tbl <- tbl %>%
          add_p() %>%
          bold_p()
      }
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

  # Résumé des tableaux pour l'export
  output$tables_summary <- renderUI({
    tables <- all_tables()
    if (length(tables) == 0) {
      return(p("Aucun tableau généré pour le moment.", class = "error-message"))
    }
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

  # Export HTML
  output$download_html <- downloadHandler(
    filename = function() {
      paste0("tableau_gtsummary_", Sys.Date(), ".html")
    },
    content = function(file) {
      req(current_table())
      html_content <- paste0(
        "<!DOCTYPE html>",
        "<html><head>",
        "<title>", input$report_title, "</title>",
        "<meta charset='UTF-8'>",
        "<style>body { font-family: Arial, sans-serif; margin: 20px; }</style>",
        "</head><body>",
        "<h1>", input$report_title, "</h1>",
        "<p>Généré le: ", Sys.Date(), "</p>",
        current_table() %>% as_gt() %>% as_raw_html(),
        "</body></html>"
      )
      writeLines(html_content, file)
    }
  )

  # Export CSV
  output$download_csv <- downloadHandler(
    filename = function() {
      paste0("donnees_", Sys.Date(), ".csv")
    },
    content = function(file) {
      req(current_data())
      write_csv(current_data(), file)
    }
  )

  # Export Word
  output$download_word <- downloadHandler(
    filename = function() {
      paste0("rapport_complet_", Sys.Date(), ".docx")
    },
    content = function(file) {
      tables <- all_tables()
      if (length(tables) == 0) {
        showNotification("Aucun tableau à exporter. Générez d'abord des analyses.", type = "warning")
        return()
      }
      tryCatch({
        doc <- read_docx()
        doc <- doc %>%
          body_add_par(value = input$report_title, style = "heading 1") %>%
          body_add_par(value = paste("Rapport généré le:", format(Sys.Date(), "%d/%m/%Y")), style = "Normal") %>%
          body_add_par(value = paste("Nombre de tableaux:", length(tables)), style = "Normal") %>%
          body_add_break()
        for (i in seq_along(tables)) {
          table_info <- tables[[i]]
          doc <- doc %>%
            body_add_par(value = paste("Tableau", i, ":", table_info$title), style = "heading 2") %>%
            body_add_par(value = paste("Type:", table_info$type, "| Généré le:", format(table_info$timestamp, "%d/%m/%Y %H:%M")), style = "Normal")
          ft <- table_info$table %>% as_flex_table()
          doc <- doc %>% body_add_flextable(ft)
          if (i < length(tables)) {
            doc <- doc %>% body_add_break()
          }
        }
        print(doc, target = file)
        showNotification("Rapport Word généré avec succès!", type = "message")
      }, error = function(e) {
        showNotification(paste("Erreur lors de la génération du rapport:", e$message), type = "error")
      })
    }
  )
}

# Lancement de l'application
shinyApp(ui, server)
