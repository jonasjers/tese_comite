#========================================#
# PROJETO CEMIT
#========================================#

# instalar pacotes
#install.packages("RQDA")

# carregar pacote
library(RQDA)

# executar Qualitative Data Analysis
RQDA()

exportCodings(file = "Exported Codings.html", Fid = NULL,
              order = c("fname", "ftime", "ctime"), append = FALSE,
              codingTable="coding")
