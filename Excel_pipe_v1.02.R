#Carga librerías e inicializa estructuras

library(readxl)
W <- list()
fecha.fin <- as.Date("2017/01/01")
fecha.origen <- as.Date("2015/12/31")
i <- 1

# Carga la excel
mi.fichero <- paste("C:/Users/a212857/Downloads/W", sprintf("%02d", i), ".xlsx", sep="")
W[[i]] <- read_excel(mi.fichero, 1)

# Obtiene el vector lógico para consultar las oportunidades con Closing date 2016
lfecha <-  as.Date.numeric(W[[i]]$`Closing Date`, origin="1899-12-30") < fecha.fin & as.Date.numeric(W[[i]]$`Closing Date`, origin="1899-12-30") > fecha.origen

# Selecciona los Status correspondientes a Oportunidades Active, Preferred Supplier y Unqualified
lstatus1 <- W[[i]]$`Status` == "Active"
lstatus2 <- W[[i]]$`Status` == "Preferred Supplier" 
lstatus3 <- W[[i]]$`Status` == "Unqualified"
lstatus <- as.vector(lstatus1 | lstatus2 | lstatus3)

#Compone la seleccion múltiple
lseleccion <- as.vector(lfecha & lstatus) 

#Obtiene la suma de OE de las oportunidades con closing date 2016
sum(W[[i]]$`Order Entry Total *1000`[lseleccion])





