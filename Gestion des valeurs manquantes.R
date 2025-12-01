# ===============================================
#  reproNA v2.0 : Application Shiny Professionnelle
#  Design moderne pour analyse épidémiologique
#  Auteur : Raider's Universe + Claude
#  Date   : 2025-10-21
# ===============================================

library(shiny)
library(DT)
library(ggplot2)
library(mice)
library(dplyr)
library(VIM)
library(gridExtra)
library(rmarkdown)
library(shinyWidgets)
library(shinydashboard)
library(bslib)

# CONFIGURATION DU THÈME PROFESSIONNEL ----------------------------------------
custom_theme <- bs_theme(
  version = 5,
  bg = "#FFFFFF",
  fg = "#2C3E50",
  primary = "#1E3A8A",      # Bleu profond
  secondary = "#64748B",    # Gris moderne
  success = "#10B981",      # Vert accent
  warning = "#F59E0B",      # Jaune accent
  danger = "#EF4444",
  base_font = font_google("Inter"),
  heading_font = font_google("Poppins"),
  font_scale = 0.95
)

# CSS PERSONNALISÉ ------------------------------------------------------------
custom_css <- "
/* === STYLE GÉNÉRAL === */
body {
  background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
  font-family: 'Inter', sans-serif;
}

.main-header {
  background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%);
  color: white;
  padding: 20px;
  border-radius: 12px;
  margin-bottom: 30px;
  box-shadow: 0 8px 16px rgba(30, 58, 138, 0.2);
}

.main-header h1 {
  font-family: 'Poppins', sans-serif;
  font-weight: 700;
  font-size: 2.5rem;
  margin: 0;
  display: flex;
  align-items: center;
  gap: 15px;
}

.logo-container {
  display: inline-block;
  background: white;
  padding: 10px;
  border-radius: 8px;
  box-shadow: 0 4px 8px rgba(0,0,0,0.1);
}

/* === SIDEBAR MODERNE === */
.well {
  background: white;
  border: none;
  border-radius: 16px;
  box-shadow: 0 4px 20px rgba(0,0,0,0.08);
  padding: 25px;
}

.sidebar-section {
  margin-bottom: 30px;
  padding-bottom: 20px;
  border-bottom: 2px solid #E5E7EB;
}

.sidebar-section:last-child {
  border-bottom: none;
}

.sidebar-section h4 {
  color: #1E3A8A;
  font-weight: 600;
  margin-bottom: 15px;
  display: flex;
  align-items: center;
  gap: 10px;
}

.section-icon {
  font-size: 1.3rem;
}

/* === BOUTONS PROFESSIONNELS === */
.btn-primary {
  background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%);
  border: none;
  border-radius: 8px;
  padding: 12px 24px;
  font-weight: 600;
  transition: all 0.3s ease;
  box-shadow: 0 4px 12px rgba(30, 58, 138, 0.3);
}

.btn-primary:hover {
  transform: translateY(-2px);
  box-shadow: 0 6px 20px rgba(30, 58, 138, 0.4);
}

.btn-success {
  background: linear-gradient(135deg, #10B981 0%, #34D399 100%);
  border: none;
  border-radius: 8px;
  font-weight: 600;
  transition: all 0.3s ease;
}

.download-btn {
  margin: 8px 0;
  width: 100%;
  border-radius: 8px;
  padding: 10px;
  font-weight: 500;
  transition: all 0.3s ease;
}

/* === ONGLETS ÉLÉGANTS === */
.nav-tabs {
  border-bottom: 2px solid #E5E7EB;
  margin-bottom: 25px;
}

.nav-tabs .nav-link {
  border: none;
  color: #64748B;
  font-weight: 600;
  padding: 12px 24px;
  margin-right: 8px;
  border-radius: 8px 8px 0 0;
  transition: all 0.3s ease;
}

.nav-tabs .nav-link:hover {
  background: #F1F5F9;
  color: #1E3A8A;
}

.nav-tabs .nav-link.active {
  background: white;
  color: #1E3A8A;
  border-bottom: 3px solid #1E3A8A;
}

/* === CARTES DE CONTENU === */
.content-card {
  background: white;
  border-radius: 12px;
  padding: 25px;
  box-shadow: 0 4px 16px rgba(0,0,0,0.08);
  margin-bottom: 20px;
}

/* === INPUTS MODERNES === */
.form-control, .form-select {
  border-radius: 8px;
  border: 2px solid #E5E7EB;
  padding: 10px 15px;
  transition: all 0.3s ease;
}

.form-control:focus, .form-select:focus {
  border-color: #3B82F6;
  box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.1);
}

/* === DATAFRAME STYLISÉ === */
.dataTables_wrapper {
  font-family: 'Inter', sans-serif;
}

.dataTables_wrapper .dataTables_length,
.dataTables_wrapper .dataTables_filter {
  margin-bottom: 15px;
}

/* === ALERTES ÉLÉGANTES === */
.alert {
  border-radius: 8px;
  border: none;
  padding: 15px 20px;
  font-weight: 500;
}

/* === STATISTIQUES EN CARTE === */
.stat-card {
  background: linear-gradient(135deg, #3B82F6 0%, #1E3A8A 100%);
  color: white;
  padding: 20px;
  border-radius: 12px;
  margin: 10px 0;
  box-shadow: 0 4px 12px rgba(30, 58, 138, 0.2);
}

.stat-value {
  font-size: 2rem;
  font-weight: 700;
}

.stat-label {
  font-size: 0.9rem;
  opacity: 0.9;
  text-transform: uppercase;
  letter-spacing: 1px;
}

/* === RESPONSIVE === */
@media (max-width: 768px) {
  .main-header h1 {
    font-size: 1.8rem;
  }
}
"

# UI MODERNE ------------------------------------------------------------------
ui <- fluidPage(
  theme = custom_theme,
  tags$head(
    tags$style(HTML(custom_css)),
    tags$link(rel = "stylesheet", 
              href = "https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css")
  ),
  
  # HEADER PROFESSIONNEL
  div(class = "main-header",
      fluidRow(
        column(12,
               h1(
                 span(class = "logo-container",
                      tags$i(class = "fas fa-database", style = "color: #1E3A8A; font-size: 2rem;")),
                 "reproNA",
                 tags$small(style = "font-size: 1rem; opacity: 0.9; font-weight: 400;",
                            "Gestion Reproductible des Valeurs Manquantes")
               )
        )
      )
  ),
  
  sidebarLayout(
    # SIDEBAR REDESIGNÉE
    sidebarPanel(
      width = 3,
      
      # SECTION 1 : Données
      div(class = "sidebar-section",
          h4(
            tags$i(class = "fas fa-upload section-icon", style = "color: #10B981;"),
            "1. Données"
          ),
          fileInput("file", NULL,
                    buttonLabel = list(tags$i(class = "fas fa-file-import"), " Parcourir"),
                    placeholder = "CSV, Excel ou RDS",
                    accept = c(".csv", ".xlsx", ".rds"))
      ),
      
      # SECTION 2 : Stratégie
      div(class = "sidebar-section",
          h4(
            tags$i(class = "fas fa-brain section-icon", style = "color: #3B82F6;"),
            "2. Stratégie"
          ),
          selectInput("strategy", "Méthode d'imputation :", 
                      choices = c("Suppression (listwise)", 
                                  "Moyenne/Médiane", 
                                  "KNN", 
                                  "MICE (pmm)"),
                      selected = "MICE (pmm)"),
          
          conditionalPanel(
            condition = "input.strategy == 'KNN'",
            numericInput("k", "k (nombre de voisins)", 5, min = 1, max = 20)
          ),
          
          conditionalPanel(
            condition = "input.strategy == 'MICE (pmm)'",
            numericInput("m_mice", "m (nb. d'imputations)", 5, min = 1, max = 50),
            numericInput("maxit", "maxit (itérations)", 10, min = 5, max = 50)
          ),
          
          numericInput("seed", "Graine aléatoire", 123, min = 1),
          
          actionButton("run", 
                       label = list(tags$i(class = "fas fa-play"), " Lancer l'analyse"),
                       class = "btn-primary btn-lg", 
                       width = "100%",
                       style = "margin-top: 15px;")
      ),
      
      # SECTION 3 : Export
      div(class = "sidebar-section",
          h4(
            tags$i(class = "fas fa-download section-icon", style = "color: #F59E0B;"),
            "3. Export"
          ),
          downloadButton("download_script", 
                         label = list(tags$i(class = "fas fa-code"), " Script R"),
                         class = "download-btn btn-outline-primary"),
          downloadButton("download_data", 
                         label = list(tags$i(class = "fas fa-table"), " Données"),
                         class = "download-btn btn-outline-success"),
          downloadButton("download_report", 
                         label = list(tags$i(class = "fas fa-file-pdf"), " Rapport"),
                         class = "download-btn btn-outline-warning")
      ),
      
      # FOOTER SIDEBAR
      tags$div(style = "text-align: center; margin-top: 20px; opacity: 0.7;",
               tags$small("Powered by", tags$br(), 
                          tags$b("Raider's Universe", style = "color: #1E3A8A;")))
    ),
    
    # MAIN PANEL MODERNISÉ
    mainPanel(
      tabsetPanel(
        id = "main_tabs",
        
        # ONGLET 1 : Diagnostic
        tabPanel(
          title = list(tags$i(class = "fas fa-chart-bar"), " Diagnostic NA"),
          value = "diag",
          br(),
          div(class = "content-card",
              h3(tags$i(class = "fas fa-microscope"), " Analyse des Valeurs Manquantes"),
              hr(),
              plotOutput("na_heatmap", height = "450px")
          ),
          div(class = "content-card",
              h4(tags$i(class = "fas fa-table"), " Résumé par Variable"),
              DT::dataTableOutput("na_summary")
          )
        ),
        
        # ONGLET 2 : Résultats
        tabPanel(
          title = list(tags$i(class = "fas fa-chart-line"), " Résultats"),
          value = "results",
          br(),
          uiOutput("stats_cards"),
          div(class = "content-card",
              h4(tags$i(class = "fas fa-exchange-alt"), " Comparaison Avant/Après"),
              plotOutput("before_after", height = "550px")
          )
        ),
        
        # ONGLET 3 : Diagnostics MICE
        tabPanel(
          title = list(tags$i(class = "fas fa-stethoscope"), " Diagnostics MICE"),
          value = "mice",
          br(),
          conditionalPanel(
            "input.strategy == 'MICE (pmm)'",
            div(class = "content-card",
                h4(tags$i(class = "fas fa-chart-area"), " Convergence"),
                plotOutput("mice_conv", height = "400px")
            ),
            div(class = "content-card",
                h4(tags$i(class = "fas fa-wave-square"), " Distribution des Imputations"),
                plotOutput("mice_density", height = "400px")
            )
          ),
          conditionalPanel(
            "input.strategy != 'MICE (pmm)'",
            div(class = "alert alert-info",
                tags$i(class = "fas fa-info-circle"),
                " Cet onglet est disponible uniquement avec la méthode MICE (pmm)."
            )
          )
        ),
        
        # ONGLET 4 : Script
        tabPanel(
          title = list(tags$i(class = "fas fa-code"), " Script"),
          value = "script",
          br(),
          div(class = "content-card",
              h4(tags$i(class = "fas fa-file-code"), " Code R Reproductible"),
              p(style = "color: #64748B;",
                "Copiez ce code pour reproduire l'analyse dans votre environnement R."),
              hr(),
              verbatimTextOutput("generated_script", placeholder = TRUE)
          )
        )
      )
    )
  ),
  
  # FOOTER GLOBAL
  tags$div(
    style = "text-align: center; padding: 30px; color: #64748B; margin-top: 50px;",
    hr(),
    p(tags$i(class = "fas fa-flask"), " reproNA v2.0 | ",
      tags$i(class = "fas fa-shield-alt"), " Analyse Reproductible | ",
      tags$i(class = "fas fa-rocket"), " Powered by Raider's Universe")
  )
)

# SERVER LOGIC ----------------------------------------------------------------
server <- function(input, output, session) {
  
  # Données brutes
  data_raw <- reactive({
    req(input$file)
    ext <- tools::file_ext(input$file$name)
    d <- switch(ext,
                csv = read.csv(input$file$datapath, stringsAsFactors = FALSE),
                xlsx = readxl::read_excel(input$file$datapath),
                rds = readRDS(input$file$datapath),
                validate("Format non supporté"))
    d
  })
  
  # Analyse principale
  result <- eventReactive(input$run, {
    set.seed(input$seed)
    d <- data_raw()
    
    showNotification("Analyse en cours...", type = "message", duration = 2)
    
    log <- list(
      date = Sys.time(),
      seed = input$seed,
      strategy = input$strategy,
      n_row_initial = nrow(d),
      n_na_initial = sum(is.na(d)),
      pct_na_initial = round(sum(is.na(d)) / (nrow(d) * ncol(d)) * 100, 2)
    )
    
    # Stratégies d'imputation
    if (input$strategy == "Suppression (listwise)") {
      d_clean <- na.omit(d)
      log$action <- "na.omit"
      log$n_row_final <- nrow(d_clean)
      log$n_suppressed <- log$n_row_initial - nrow(d_clean)
      
    } else if (input$strategy == "Moyenne/Médiane") {
      d_clean <- d
      for(col in names(d)) {
        if(is.numeric(d[[col]])) {
          miss <- is.na(d[[col]])
          if(any(miss)) {
            med <- median(d[[col]], na.rm = TRUE)
            d_clean[[col]][miss] <- med
            log[[paste0("med_", col)]] <- round(med, 3)
          }
        }
      }
      log$action <- "Imputation par médiane"
      
    } else if (input$strategy == "KNN") {
      k <- input$k
      d_clean <- VIM::kNN(d, k = k, numFun = median, catFun = VIM::maxCat)[,1:ncol(d)]
      log$action <- paste("KNN avec k =", k)
      
    } else if (input$strategy == "MICE (pmm)") {
      m <- input$m_mice
      maxit <- input$maxit
      
      meth <- make.method(d)
      meth[] <- "pmm"
      pred <- make.predictorMatrix(d)
      pred[,] <- 1; diag(pred) <- 0
      
      imp <- mice(d, m = m, maxit = maxit, method = meth, 
                  predictorMatrix = pred, seed = input$seed, printFlag = FALSE)
      
      d_clean <- complete(imp, 1)
      log$action <- paste("MICE (pmm) : m =", m, ", maxit =", maxit)
      log$mids <- imp
    }
    
    log$n_na_final <- sum(is.na(d_clean))
    
    showNotification("✓ Analyse terminée avec succès !", type = "message", duration = 3)
    
    list(data = d_clean, raw = d, log = log, strategy = input$strategy)
  })
  
  # Cartes statistiques
  output$stats_cards <- renderUI({
    req(result())
    res <- result()
    log <- res$log
    
    tagList(
      fluidRow(
        column(3,
               div(class = "stat-card",
                   div(class = "stat-value", log$n_row_initial),
                   div(class = "stat-label", "Observations initiales")
               )
        ),
        column(3,
               div(class = "stat-card",
                   div(class = "stat-value", log$n_na_initial),
                   div(class = "stat-label", "NA détectées")
               )
        ),
        column(3,
               div(class = "stat-card",
                   div(class = "stat-value", paste0(log$pct_na_initial, "%")),
                   div(class = "stat-label", "% de NA")
               )
        ),
        column(3,
               div(class = "stat-card", style = "background: linear-gradient(135deg, #10B981 0%, #059669 100%);",
                   div(class = "stat-value", log$n_na_final),
                   div(class = "stat-label", "NA finales")
               )
        )
      ),
      br(),
      div(class = "alert alert-success",
          tags$i(class = "fas fa-check-circle"),
          tags$b(" Méthode appliquée : "), log$action
      )
    )
  })
  
  # Heatmap NA
  output$na_heatmap <- renderPlot({
    req(data_raw())
    d <- data_raw()
    if(sum(is.na(d)) == 0) {
      par(mar = c(1,1,1,1))
      plot(1, type="n", axes=FALSE, xlab="", ylab="", xlim=c(0,10), ylim=c(0,10))
      text(5, 5, "✓ Aucune valeur manquante détectée", cex = 2.5, col = "#10B981", font = 2)
    } else {
      VIM::aggr(d, numbers = TRUE, prop = FALSE, sortVars = TRUE, 
                col = c("#3B82F6", "#EF4444"), 
                border = NA, cex.axis = 0.8)
    }
  })
  
  # Tableau récapitulatif NA
  output$na_summary <- DT::renderDataTable({
    req(data_raw())
    d <- data_raw()
    na_count <- sapply(d, function(x) sum(is.na(x)))
    na_pct <- round(na_count / nrow(d) * 100, 2)
    
    df <- data.frame(
      Variable = names(d),
      Type = sapply(d, function(x) class(x)[1]),
      NA_Count = na_count,
      NA_Percent = na_pct
    ) %>% 
      arrange(desc(NA_Count))
    
    datatable(df, 
              rownames = FALSE,
              options = list(
                pageLength = 15,
                dom = 'Bfrtip',
                columnDefs = list(
                  list(className = 'dt-center', targets = 1:3)
                )
              )) %>%
      formatStyle('NA_Percent',
                  background = styleColorBar(range(df$NA_Percent), '#3B82F6'),
                  backgroundSize = '90% 70%',
                  backgroundRepeat = 'no-repeat',
                  backgroundPosition = 'center')
  })
  
  # Comparaison avant/après
  output$before_after <- renderPlot({
    req(result())
    res <- result()
    d1 <- res$raw
    d2 <- res$data
    vars <- names(d1)[sapply(d1, is.numeric)]
    
    if(length(vars) == 0) {
      par(mar = c(1,1,1,1))
      plot(1, type="n", axes=FALSE, xlab="", ylab="")
      text(1, 1, "Aucune variable numérique à visualiser", cex = 1.5, col = "#64748B")
      return()
    }
    
    p1 <- ggplot(stack(d1[vars]), aes(values)) + 
      geom_density(alpha = 0.6, fill = "#EF4444", color = "#DC2626", size = 1) + 
      facet_wrap(~ind, scales = "free") + 
      ggtitle("Avant imputation") +
      theme_minimal(base_size = 12) +
      theme(
        plot.title = element_text(face = "bold", size = 14, color = "#2C3E50"),
        strip.text = element_text(face = "bold", color = "#1E3A8A"),
        panel.grid.minor = element_blank()
      )
    
    p2 <- ggplot(stack(d2[vars]), aes(values)) + 
      geom_density(alpha = 0.6, fill = "#3B82F6", color = "#1E40AF", size = 1) + 
      facet_wrap(~ind, scales = "free") + 
      ggtitle("Après imputation") +
      theme_minimal(base_size = 12) +
      theme(
        plot.title = element_text(face = "bold", size = 14, color = "#2C3E50"),
        strip.text = element_text(face = "bold", color = "#1E3A8A"),
        panel.grid.minor = element_blank()
      )
    
    grid.arrange(p1, p2, ncol = 2)
  })
  
  # Diagnostics MICE
  output$mice_conv <- renderPlot({
    req(result()$strategy == "MICE (pmm)")
    plot(result()$log$mids, main = "Convergence des paramètres MICE")
  })
  
  output$mice_density <- renderPlot({
    req(result()$strategy == "MICE (pmm)")
    densityplot(result()$log$mids)
  })
  
  # Script généré
  output$generated_script <- renderText({
    req(result())
    res <- result()
    
    script <- paste0(
      "# ============================================\n",
      "#  reproNA : Script Reproductible\n",
      "#  Généré le : ", format(Sys.time(), "%Y-%m-%d %H:%M:%S"), "\n",
      "#  Méthode   : ", res$log$action, "\n",
      "#  Seed      : ", input$seed, "\n",
      "# ============================================\n\n",
      "set.seed(", input$seed, ")\n\n",
      "# Charger les données\n",
      "# data <- read.csv('votre_fichier.csv')\n\n"
    )
    
    if(res$strategy == "Suppression (listwise)") {
      script <- paste0(script, 
                       "# Suppression des lignes avec NA\n",
                       "data_clean <- na.omit(data)\n",
                       "cat('Lignes supprimées :', nrow(data) - nrow(data_clean), '\\n')\n")
      
    } else if(res$strategy == "Moyenne/Médiane") {
      script <- paste0(script, "# Imputation par médiane\n")
      for(col in names(res$raw)) {
        if(is.numeric(res$raw[[col]])) {
          miss <- is.na(res$raw[[col]])
          if(any(miss)) {
            med <- median(res$raw[[col]], na.rm = TRUE)
            script <- paste0(script, "data$", col, "[is.na(data$", col, ")] <- ", 
                             round(med, 3), "  # Médiane\n")
          }
        }
      }
      script <- paste0(script, "\ndata_clean <- data\n")
      
    } else if(res$strategy == "KNN") {
      script <- paste0(script,
                       "library(VIM)\n\n",
                       "# Imputation KNN\n",
                       "data_clean <- kNN(data, k = ", input$k, 
                       ", numFun = median)[, 1:ncol(data)]\n")
      
    } else if(res$strategy == "MICE (pmm)") {
      script <- paste0(script,
                       "library(mice)\n\n",
                       "# Configuration MICE\n",
                       "imp <- mice(data, \n",
                       "            m = ", input$m_mice, ",\n",
                       "            maxit = ", input$maxit, ",\n",
                       "            method = 'pmm',\n",
                       "            seed = ", input$seed, ",\n",
                       "            printFlag = FALSE)\n\n",
                       "# Récupérer les données imputées (1ère imputation)\n",
                       "data_clean <- complete(imp, 1)\n\n",
                       "# Diagnostics\n",
                       "plot(imp)  # Convergence\n",
                       "densityplot(imp)  # Distributions\n")
    }
    
    script <- paste0(script,
                     "\n# Vérification\n",
                     "cat('NA restantes :', sum(is.na(data_clean)), '\\n')\n")
    
    script
  })
  
  # Téléchargements
  output$download_script <- downloadHandler(
    filename = function() paste0("reproNA_script_", Sys.Date(), ".R"),
    content = function(file) {
      writeLines(output$generated_script(), file)
    }
  )
  
  output$download_data <- downloadHandler(
    filename = function() paste0("data_imputed_", Sys.Date(), ".csv"),
    content = function(file) {
      write.csv(result()$data, file, row.names = FALSE)
    }
  )
  
  output$download_report <- downloadHandler(
    filename = function() paste0("reproNA_rapport_", Sys.Date(), ".html"),
    content = function(file) {
      res <- result()
      log <- res$log
      
      rmd_content <- paste0("---\n",
                            "title: 'reproNA - Rapport d\\'Analyse'\n",
                            "date: '", Sys.Date(), "'\n",
                            "output: \n",
                            "  html_document:\n",
                            "    theme: cosmo\n",
                            "    toc: true\n",
                            "    toc_float: true\n",
                            "---\n\n",
                            "## Résumé Exécutif\n\n",
                            "- **Méthode** : ", log$action, "\n",
                            "- **Observations initiales** : ", log$n_row_initial, "\n",
                            "- **NA initiales** : ", log$n_na_initial, 
                            " (", log$pct_na_initial, "%)\n",
                            "- **Seed** : ", log$seed, "\n\n",
                            "## Conclusion\n\n",
                            "Analyse reproductible générée par reproNA v2.0.\n")
      
      temp_rmd <- tempfile(fileext = ".Rmd")
      writeLines(rmd_content, temp_rmd)
      rmarkdown::render(temp_rmd, output_file = file, quiet = TRUE)
    }
  )
}
6
# LANCEMENT -------------------------------------------------------------------
shinyApp(ui = ui, server = server)