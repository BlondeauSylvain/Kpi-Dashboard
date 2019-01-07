rm(list = ls())

library(rJava)
library(xlsxjars)
library(xlsx)
library(tidyr)
library(tidyverse)
library(stringr)
library(plyr)
library(dplyr)

#chargement data 
data<-read.xlsx("./sx.xlsx",sheetIndex = 1,colIndex=c(1:5,17:18,20:22),header=TRUE,stringsAsFactors=FALSE,encoding="UTF-8",dec=",")
data<-as.data.frame(data[-1,])

for(i in 1:nrow(data)){
  x<-substring(data$Numero.d.affaire[i],1,5)  
  data$Agences[i]<-x}

for(i in 1:nrow(data)){
  if (data$Agences[i] == "unitp") data$Agences[i]<-"BUp"
  if (data$Agences[i] == "unitb") data$Agences[i]<-"BUb"
  if (data$Agences[i] =="unitc") data$Agences[i]<-"BUc"
  if (data$Agences[i] =="unitd") data$Agences[i]<-"BUd" 
  if (data$Agences[i] =="unite") data$Agences[i]<-"BUe"
  if (data$Agences[i] =="unitf") data$Agences[i]<-"BUf"
  if (data$Agences[i] =="unitg") data$Agences[i]<-"BUg"
  } 

#################################################################################################
############################UI & SERVER ##########################
library(shiny)
require(shinydashboard)
library(ggplot2)
library(RColorBrewer)
library(extrafont)
library(data.table)

ui <- dashboardPage(
  dashboardHeader(title = "KPI"),
  dashboardSidebar(
    sidebarSearchForm(textId = "searchText", buttonId = "searchButton", label = "Search..."),
    sidebarMenu(
      menuItem(p(h4("Dataset")), tabName = "data")
    )
  ),
  
 dashboardBody(
    tabItems(
      tabItem(tabName ="data" ,
              fluidRow(
                tabsetPanel(type = "tabs",
                            tabPanel("Dataset",
                                    fluidRow(
                                      box(title = "Data Viewer"
                                         ,width = 12
                                         ,id = "dataTabBox"
                                         ,column(4,selectInput("num","Numéro Projet",choices=c("All", str_trim(unique(data$Numero.d.affaire)),selectize = TRUE)))
                                         ,column(4,selectInput("client","Client:",choices=c("All", str_trim(unique(data$Client)),selectize = TRUE)))
                                         ,column(4,selectInput("statut","Statut:",choices=c("All", unique(data$Statut)),selectize = TRUE))
                                         ,column(4,selectInput("agences","Agences:",choices=c("All", unique(data$Agences)),selectize = TRUE))
                                         ,column(4,selectInput("responsable"," Responsable:",choices=c("All", str_trim(unique(data$Responsable.projet)),selectize = TRUE)))
                                         )
                                      ,box(title = "dataset"
                                            ,width = 12
                                            ,id = "dataTabBox"
                                            ,dataTableOutput("sx")
                                          )
                                        )
                                  ),
                            tabPanel("Indicateurs Projets",
                                     fluidRow(
                                        box(title = "Performance"
                                         ,width = 12
                                         ,id = "dataTabBox"
                                         ,valueBoxOutput("performance")
                                         ,valueBoxOutput("heures")
                                         ,valueBoxOutput("nombreprojet")
                                         ,br(),br(),br(),br(),br(),br(),br()
                                         ,hr()
                                         ,verbatimTextOutput("txt")
                                        )
                                        ),
                                    fluidRow(
                                      box(title = "Heures commandées et engagées par tranche d'heures"
                                           ,status = "primary"
                                           ,solidHeader = TRUE
                                           ,collapsible = TRUE
                                           ,width = 6
                                           ,id = "dataTabBox"
                                           ,plotOutput("graph",height = "300px")
                                      )
                                      ,box(title = "Performance"
                                           ,status = "primary"
                                           ,solidHeader = TRUE
                                           ,collapsible = TRUE
                                           ,width = 6
                                           ,id = "dataTabBox"
                                           ,plotOutput("graph_perf",height = "300px")
                                      )
                                      ,box(title = "Synthèse KPI"
                                          ,status = "primary"
                                          ,solidHeader = TRUE
                                          ,collapsible = TRUE
                                          ,width = 12
                                          ,height = 590
                                          ,id = "dataTabBox"
                                          ,dataTableOutput("tableau")
                                           ,hr()
                                           ,column(4,numericInput("pas", "Pas",value = 200))
                                          )
                                  
                                      ,column(12,
                                         box(title = "Generate report"
                                        ,status = "primary"
                                          ,solidHeader = TRUE
                                          ,collapsible = TRUE
                                          ,width = 6
                                          ,id = "dataTabBox"
                                         ,textInput("pdfname", "Filename", "My.pdf")
                                         ,downloadButton("outputButton", "Download PDF")
                                        
                                         ))
                                        )
                                    )
                                )
                          )
                )
        )
  )
                
)
# create the server functions for the dashboard  
    server <- function(input, output)  {
      
      #  data table "Données SX","
      output$sx <-  renderDataTable({
        if (input$num != "All") {data <- data[str_trim(data$Numero.d.affaire) == input$num,]}
        if (input$client != "All") {data <- data[data$Client== input$client,]}
        if (input$statut != "All") {data <- data[data$Statut == input$statut,]}
        if (input$agences != "All") {data <- data[data$Agences == input$agences,]}
        if (input$responsable != "All") {data <- data[str_trim(data$Responsable.projet) == input$responsable,]}
        data
        })
      
      tdonne<-reactive({
        tab<-data.frame(pas = c(1:6,0), Heures_commandes =c(1:6,0), Heures_engagees=c(1:6,0), RAF=c(1:6,0),Bilan_heures = c(1:6,0), Performance = c(1:6,0), Nombre_projet = c(1:6,0))
        for (i in 1:nrow(tab))
          for(j in 1:ncol(tab)){tab[i,j]<-0}
        
        for (j in 1:nrow(tab)){
          tab$pas[j]<-input$pas*j  #variable entrée du pas
          nb=0
            if((j-1) == 0) {var = 0 }
            else{var = tab$pas[j-1]}
            var_max <-0
            #var<-input$nbligne+1
            if(j == 7) {var_max = max(data$Commande.heures)}
            else{var_max = tab$pas[j]}
            #BOUCLE Filtre pour statut, agences, client
            for (i in 1:nrow(data)){
                if (data$Commande.heures[i]>var && data$Commande.heures[i]<=var_max && input$statut == "All" && input$agences == "All" && input$client == "All" && input$responsable =="All"|
                    data$Commande.heures[i]>var && data$Commande.heures[i]<=var_max && str_trim(data$Statut[i])==input$statut && input$agences == "All" && input$client == "All" && input$responsable =="All"|
                    data$Commande.heures[i]>var && data$Commande.heures[i]<=var_max && str_trim(data$Agences[i])==input$agences && input$statut == "All"  && input$client == "All" && input$responsable =="All"|
                    data$Commande.heures[i]>var && data$Commande.heures[i]<=var_max && str_trim(data$Statut[i])==input$statut && str_trim(data$Agences[i])==input$agences && input$client == "All" && input$responsable =="All"|
                  
                    data$Commande.heures[i]>var && data$Commande.heures[i]<=var_max && input$statut == "All" && input$agences == "All" && str_trim(data$Client[i])==input$client && input$responsable =="All" |
                    data$Commande.heures[i]>var && data$Commande.heures[i]<=var_max && str_trim(data$Statut[i])==input$statut && input$agences == "All" && str_trim(data$Client[i])==input$client && input$responsable =="All" |
                    data$Commande.heures[i]>var && data$Commande.heures[i]<=var_max && str_trim(data$Agences[i])==input$agences && input$statut == "All"  && str_trim(data$Client[i])==input$client && input$responsable =="All"|
                    data$Commande.heures[i]>var && data$Commande.heures[i]<=var_max && str_trim(data$Statut[i])==input$statut && str_trim(data$Agences[i])==input$agences && str_trim(data$Client[i])==input$client && input$responsable =="All"|
                  
                    data$Commande.heures[i]>var && data$Commande.heures[i]<=var_max && input$statut == "All" && input$agences == "All" && input$client == "All" && str_trim(data$Responsable.projet[i])==input$responsable|
                    data$Commande.heures[i]>var && data$Commande.heures[i]<=var_max && str_trim(data$Statut[i])==input$statut && input$agences == "All" && input$client == "All" && str_trim(data$Responsable.projet[i])==input$responsable|
                    data$Commande.heures[i]>var && data$Commande.heures[i]<=var_max && str_trim(data$Agences[i])==input$agences && input$statut == "All"  && input$client == "All" && str_trim(data$Responsable.projet[i])==input$responsable|
                    data$Commande.heures[i]>var && data$Commande.heures[i]<=var_max && str_trim(data$Statut[i])==input$statut && str_trim(data$Agences[i])==input$agences && input$client == "All" && str_trim(data$Responsable.projet[i])==input$responsable|
                  
                    data$Commande.heures[i]>var && data$Commande.heures[i]<=var_max && input$statut == "All" && input$agences == "All" && str_trim(data$Client[i])==input$client && str_trim(data$Responsable.projet[i])==input$responsable |
                    data$Commande.heures[i]>var && data$Commande.heures[i]<=var_max && str_trim(data$Statut[i])==input$statut && input$agences == "All" && str_trim(data$Client[i])==input$client && str_trim(data$Responsable.projet[i])==input$responsable |
                    data$Commande.heures[i]>var && data$Commande.heures[i]<=var_max && str_trim(data$Agences[i])==input$agences && input$statut == "All"  && str_trim(data$Client[i])==input$client && str_trim(data$Responsable.projet[i])==input$responsable|
                    data$Commande.heures[i]>var && data$Commande.heures[i]<=var_max && str_trim(data$Statut[i])==input$statut && str_trim(data$Agences[i])==input$agences && str_trim(data$Client[i])==input$client && str_trim(data$Responsable.projet[i])==input$responsable)
                {
                tab$Heures_commandes[j]<-tab$Heures_commandes[j] + data$Commande.heures[i]
                tab$Heures_engagees[j]<-tab$Heures_engagees[j] + data$Realise.heures[i]
                tab$RAF[j]<-tab$RAF[j] + data$Reste.a.faire[i]
                tab$Bilan_heures[j]  <-tab$Bilan_heures[j] + data$Bilan.heures[i]
                tab$Performance[j]  <-round((tab$Bilan_heures[j] / tab$Heures_commandes[j])*100,digit = 1) #,"%")
                 # tab$Performance[j]  <-paste(round((tab$Bilan_heures[j] / tab$Heures_commandes[j])*100,digit = 1) #,"%")                                            
                nb<-nb+1
                tab$Nombre_projet[j]  <- nb}
            }
             if (j==7) tab$pas[j]<-paste(">",6*input$pas) 
        }
        tab
        })
        
      output$tableau <- renderDataTable({tdonne()})                            # Creation du tableau synthèse

      output$performance<-renderValueBox({                                     # Valeur performance projet en fonction du tableau synthèse
          perf<-tdonne()
          perf.project<-(sum(perf$Bilan_heures)/sum(perf$Heures_commandes))*100
          valueBox(
            paste0(round(perf.project, digits = 1),"%")      #perf.project
            ,"Performance"
            ,icon = icon("pie-chart")
            ,color = ifelse(perf.project < -4, "red", 
                            ifelse(perf.project< 0, "yellow","green"))
  
          )
      })
      
      output$heures <-renderValueBox({                                          #Heures projet en fonction du tableau synthèse
        h.project<-tdonne()
        heure.project<-sum(h.project$Heures_commandes)
        valueBox(
          paste0(round(heure.project, digits=0), "h")
          ,"Heures commandées"
          ,icon = icon("eur")
          ,color = "aqua"
        )
      })
      
      output$nombreprojet <-renderValueBox({                                           #Nombre de projet en fonction du tableau synthèse
        nbre.project<-tdonne()
        nombre.project<-sum(nbre.project$Nombre_projet)
        valueBox(
          paste0(round(nombre.project, digits=0))
          ,"Nombre de Projet"
          #,icon = icon("eur")
          ,color = "aqua"
        )
      })
  
    output$txt<-renderText({paste(Client = input$client,input$agences,input$statut,input$responsable , sep=",")})
    
    output$graph<- renderPlot({
      kpi.project<-tdonne()
  
      #creation de la version tidy de kpi.project heures engagées & RAF
      kpi.project_1.tdy<-kpi.project %>%
        gather("types","heures",3:4)
      kpi.project_1.tdy<-as.data.frame(kpi.project_1.tdy)
     
      #creation de la version tidy de kpi.project heures commandées
      kpi.project_2.tdy<-kpi.project %>%
        gather("types","heures",2:2)
      kpi.project_2.tdy<-as.data.frame(kpi.project_2.tdy)
     
      var_1<-as.numeric(factor(kpi.project_1.tdy$pas, levels = (unique(kpi.project_1.tdy$pas)), ordered=FALSE))
      var_2<-as.numeric(factor(kpi.project_2.tdy$pas, levels = (unique(kpi.project_2.tdy$pas)), ordered=FALSE))
      
      barwidth = 0.35
      ggplot() + 
          geom_bar(data = kpi.project_2.tdy,
                 mapping=  aes(x=var_2 , y=heures,fill=types),
                 stat ="identity", 
                 position = "stack",
                 width = barwidth) +
          geom_bar(data = kpi.project_1.tdy,
                 mapping=  aes(x=var_1 + barwidth + 0.01, y=heures,fill=types),
                 stat ="identity", 
                 position = "stack",
                 width = barwidth) +
        scale_fill_manual(values = c("#66CC66","#56B4E9", "#CC79A7")) +
        xlab("pas") + ylab("heures") +
        labs(fill = "Type d'heure")+
        theme(
          panel.background = element_blank(),
          panel.grid.minor = element_blank(), 
          panel.grid.major = element_line(color = "gray50", size = 0.5),
          panel.grid.major.x = element_blank(),
          axis.text.y = element_text(colour="#68382C", size = 9))
      
    })
   
    output$graph_perf<- renderPlot({
      perf.project<-tdonne()
      perf.project.tdy <- perf.project %>%
        gather("types", "perf", 6:6) %>%
        # Create the variable you need for the plot
        mutate(pas = factor(pas, levels = unique(pas), ordered=TRUE),
               fill_col = case_when(
                 perf > 0   ~ "green",
                 perf <= -4 ~ "red",
                 TRUE       ~ "orange"
               ))
      
      barwidth= 0.8
      ggplot(data = perf.project.tdy, aes(x=pas, y=perf)) +
        geom_col(aes(fill= fill_col),
                 width = barwidth) +
        geom_text(aes(label=paste(perf, "%")), vjust=1.6, color="blue", size=3.5) +
        scale_fill_identity(guide = "legend") 
 
      })
    
    output$outputButton <- downloadHandler(input$pdfname, function(theFile) {
      makePdf()
    })
    
    
    # Sample pdf-generating function:
    makePdf <- function(){
      pdf(file = "./myGenerated.pdf")
      plot()
      dev.off()
    }
    
    
}
shinyApp(ui, server) 


    