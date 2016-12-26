#Carga librerias 
library(readxl)
library(xlsx)
library(ggplot2)

# Inicializa estructuras
W <- list()
fecha.desde <- as.Date("2015/12/31")
fecha.hasta <- as.Date("2017/01/01")

Mi.Consulta <- function (WX){
  
   # Obtiene el vector logico para consultar las oportunidades con Closing date 2016
  lfechaDesde <- as.Date.numeric(WX$`Closing Date`, origin="1899-12-30") > fecha.desde
  lfechaHasta <- as.Date.numeric(WX$`Closing Date`, origin="1899-12-30") < fecha.hasta
  lfecha <- lfechaDesde & lfechaHasta 

  # Selecciona los Status correspondientes a Oportunidades Active, Preferred Supplier y Unqualified
  #lstatus <- WX$`Status` == "Active" | WX$`Status` == "Preferred Supplier" | WX$`Status` == "Unqualified"
  lstatus <- WX$`Status` == "Won"

  # Selecciona el Item Offering
  loffering <- WX$`Item Offering` == "BDS- BID Big Data Appliances" | WX$`Item Offering` == "BDS- BID bullion Servers" | WX$`Item Offering` == "MS- Products - bullion Servers"
  #loffering <- WX$`Item Offering` == "BDS- CYS Cyber Security IAM"
  
  #Compone la seleccion multiple
  lseleccion <-lstatus & lfecha
  
  
  #Obtiene la suma de OE de las oportunidades con las condiciones de Status y Closing dates dadas
  return(sum(WX$`Order Entry Total *1000`[lseleccion]))
}


for (i in 1:50){
    # Carga la excel
    W[[i]] <- read_excel(paste("C:/Users/a212857/Downloads/W", sprintf("%02d", i), ".xlsx", sep=""))
}

# Aplica la selección común a todos los data-frames y deja cada resultado en una posición del vector
P <- sapply(W, Mi.Consulta)

# Imprime en consola y en fichero "P.xlxs"
print(P)
qplot(1:length(P),P, geom="line")
write.xlsx(P, file="C:/Users/a212857/Downloads/P.xlsx")






