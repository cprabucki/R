#Carga librerias
library(readxl)
library(xlsx)
library(ggplot2)

# Inicializa estructuras
# Asume que la lista "W" con los pipe semanales se encuentra disponible en el entorno
# Para cargar W, ejecutar el script Excel_pipe_loadvX.XX.R 
fecha.desde <- as.Date("2015/12/31")
fecha.hasta <- as.Date("2017/01/01")


Mi.Consulta <- function (WX){
  
  #Inicializa los vectores a TRUE para que puedan modificar las consultas 'OR' del final
  lfecha <-    !logical(length(WX))
  loffering <- !logical(length(WX))
  lstatus <-   !logical(length(WX))
  
       # Obtiene el vector logico para consultar las oportunidades con Closing date 2016
  #lfechaDesde <- as.Date.numeric(WX$`Closing Date`, origin="1899-12-30") > fecha.desde
  #lfechaHasta <- as.Date.numeric(WX$`Closing Date`, origin="1899-12-30") < fecha.hasta
  #lfecha <- lfechaDesde & lfechaHasta 

  # Selecciona los Status correspondientes a Oportunidades Active, Preferred Supplier y Unqualified
  #lstatus <- WX$`Status` == "Active" | WX$`Status` == "Preferred Supplier" | WX$`Status` == "Unqualified"
  #lstatus <- WX$`Status` == "Won"
  lstatus <- !WX$`Status` == "Unqualified"

  # Selecciona el Item Offering
  #loffering <- WX$`Item Offering` == "BDS- BID Big Data Appliances" | WX$`Item Offering` == "BDS- BID bullion Servers" | WX$`Item Offering` == "MS- Products - bullion Servers"
  #loffering <- WX$`Item Offering` == "BDS- CYS Cyber Security IAM"
  
  #Compone la seleccion multiple
  lseleccion <-lstatus & lfecha & loffering

  #Obtiene la suma de OE de las oportunidades con las condiciones de Status y Closing dates dadas
  return(sum(WX$`Order Entry Total *1000`[lseleccion]))
}

# Aplica la selección común a todos los data-frames y deja cada resultado en una posición del vector
P <- sapply(W, Mi.Consulta)

# Imprime en consola y en fichero "P.xlxs"
print(P)
#qplot(1:length(P),P, geom="line")
#write.xlsx(P, file="C:/Users/a212857/Downloads/P.xlsx")






