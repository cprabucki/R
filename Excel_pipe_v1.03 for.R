#Carga librerias 
library(readxl)
library(xlsx)

# Inicializa estructuras
fecha.fin <- as.Date("2017/01/01")
fecha.origen <- as.Date("2015/12/31")
W <- list()
P <- c()

for (i in 1:48){
    # Carga la excel
    W[[i]] <- read_excel(paste("C:/Users/a212857/Downloads/W", sprintf("%02d", i), ".xlsx", sep=""))

    # Obtiene el vector lógico para consultar las oportunidades con Closing date 2016
    lfecha <-  as.Date.numeric(W[[i]]$`Closing Date`, origin="1899-12-30") < fecha.fin & as.Date.numeric(W[[i]]$`Closing Date`, origin="1899-12-30") > fecha.origen

    # Selecciona los Status correspondientes a Oportunidades Active, Preferred Supplier y Unqualified
    lstatus <- W[[i]]$`Status` == "Active" | W[[i]]$`Status` == "Preferred Supplier" | W[[i]]$`Status` == "Unqualified"

    #Compone la seleccion múltiple
    lseleccion <-lfecha & lstatus 

    #Obtiene la suma de OE de las oportunidades con las condiciones de Status y Closing dates dadas
    P[[i]] <- sum(W[[i]]$`Order Entry Total *1000`[lstatus])
}

# Imprime en consola y en fichero "P.xlxs"
print(P)
write.xlsx(P, file="C:/Users/a212857/Downloads/P.xlsx")






