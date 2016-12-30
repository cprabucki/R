
carga <- function (){

  #Carga librerias
  library(readxl)
  library(xlsx)

  # Inicializa estructuras
  mypath <- file.path(getwd(),"W", "W")
  W <- list()

  # Carga en memoria la megaestructura "W" como lista de datasets (pipe), uno por semana
  for (i in 1:52){
    # Carga la excel
    W[[i]] <- read_excel(paste(mypath, sprintf("%02d", i), ".xlsx", sep=""))
  }
  return(W)
}