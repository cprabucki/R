
analiza <- function(W){

    #Carga librerias
    library(readxl)
    library(xlsx)
    library(ggplot2)
    
    # Inicializa estructuras
    # Asume que la lista "W" con los pipe semanales se encuentra disponible en el entorno
    # Para cargar W, ejecutar el script Excel_pipe_loadvX.XX.R 
    fecha.desde <- as.Date("2015/12/31")
    fecha.hasta <- as.Date("2017/01/01")
    
    miConsulta <- function(Z){
      
      #Inicializa los vectores a TRUE para que puedan modificar las consultas 'OR' del final
      lfecha <-    !logical(length(Z))
      loffering <- !logical(length(Z))
      lstatus <-   !logical(length(Z))
      
      # Obtiene el vector logico para consultar las oportunidades con Closing date 2016
      #lfechaDesde <- as.Date.numeric(Z$`Closing Date`, origin="1899-12-30") > fecha.desde
      #lfechaHasta <- as.Date.numeric(Z$`Closing Date`, origin="1899-12-30") < fecha.hasta
      #lfecha <- lfechaDesde & lfechaHasta 
    
      # Selecciona los Status correspondientes a Oportunidades Active, Preferred Supplier y Unqualified
      #lstatus <- Z$`Status` == "Active" | Z$`Status` == "Preferred Supplier" | Z$`Status` == "Unqualified"
      #lstatus <- Z$`Status` == "Won"
      lstatus <- Z$`Status` == "Unqualified"
    
      # Selecciona el Item Offering
      #loffering <- Z$`Item Offering` == "BDS- BID Big Data Appliances" | Z$`Item Offering` == "BDS- BID bullion Servers" | Z$`Item Offering` == "MS- Products - bullion Servers"
      #loffering <- Z$`Item Offering` == "BDS- CYS Cyber Security IAM"
      
      #Compone la seleccion multiple
      lseleccion <-lstatus & lfecha & loffering
    
      #Obtiene la suma de OE de las oportunidades con las condiciones de Status y Closing dates dadas
      return(sum(Z$`Order Entry Total *1000`[lseleccion]))
      #return(sum(Z$`Weighted Order Entry Total *1000`[lseleccion]))
      #return(sum(Z$`Weighted Order Entry Total *1000`[lseleccion])/sum(Z$`Order Entry Total *1000`[lseleccion]))
    }
    
    # Aplica la selección común a todos los data-frames y deja cada resultado en una posición del vector
    P <- sapply(W, miConsulta)
    
    # Imprime en consola y en fichero "P.xlxs"
    write.xlsx(P, file=file.path(getwd(),"W","P.xlsx"))
    return(P)
    
    #qplot(1:length(P),P, geom="line")
}





