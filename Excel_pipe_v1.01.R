# Carga la excel
library(readxl)

W01 <- read_excel("C:/Users/a212857/Downloads/W01.xlsx", 1)
# Obtiene el vector lógico para consultar las oportunidades con Closing date 2016
lfecha.closing <- as.Date.numeric(W01$`Closing Date`, origin="1899-12-30") < as.Date("2017/01/01") & as.Date.numeric(W01$`Closing Date`, origin="1899-12-30") > as.Date("2015/12/31")
#Obtiene la suma de OE de las oportunidades con closing date 2016
sum(W01$`Order Entry Total *1000`[lfecha.closing])

W <- list(W01)
lfecha.closing.vector <- list(lfecha.closing)


W47 <- read_excel("C:/Users/a212857/Downloads/W47.xlsx", 1)
# Obtiene el vector lógico para consultar las oportunidades con Closing date 2016
lfecha.closing <- as.Date.numeric(W47$`Closing Date`, origin="1899-12-30") < as.Date("2017/01/01") & as.Date.numeric(W47$`Closing Date`, origin="1899-12-30") > as.Date("2015/12/31")
#Obtiene la suma de OE de las oportunidades con closing date 2016
sum(W47$`Order Entry Total *1000`[lfecha.closing])

W[[length(W)+1]] <- W47
lfecha.closing.vector[[length(lfecha.closing.vector)+1]] <- lfecha.closing


