# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #
# titulo: Processamento de arquivo de saida do equipamento Elisa de Varredura #
# autor:  Eula Carrara <eulacarrara@gmail.com>                                #
# v1:     03-Set-2025                                                         #
# v2:     25-Nov-2025                                                         #
# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #

## # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #
## Serao gerados 5 arquivos finais:                                                                          #
## resultados.xlsx  => resultados em uma coluna apenas                                                       #
## resultados_colunas.xlsx => resultados em colunas, para os dois comprimentos de onda juntos                #
## resultados_colunas_compr_onda_1.xlsx => resultados em colunas, apenas para o comprimento de onda 1        #
## resultados_colunas_compr_onda_2.xlsx => resultados em colunas, apenas para o comprimento de onda 2       ##
## resultados_transposto.xlsx => resultados_colunas.xlsx, mas com os comprimentos de onda na primeira linha  #
## # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #


# # # #
rm(list = ls())

# Defina o caminho completo de onde esta o seu arquivo (use \\ ou /)
dir <- "/Users/eulacarrara/Library/CloudStorage/OneDrive-Personal/Sharing/Mateus_UCP/v2"
# Defina o nome do seu arquivo com os resultados (tem que estar em .xlsx; xls nao eh aceito)
xlsx <- "Plate_Sacarose.xlsx"
# Insira o número de comprimentos de onda (nesse caso, 3)
comp_ondas_pre_def <- c(1,2,3)

# # # #

# Selecione todo os comandos abaixo e rode de uma vez so
# INICIO
# Instalar e carregar pacotes exigidos
pacotes <- c("readxl", "tidyverse", "gtools", "openxlsx")
for(p in pacotes){
  if(!require(p, character.only = TRUE, quietly = TRUE)){
    message(paste("Pacote", p, "nao instalado. Instalando..."))
    install.packages(p)
    if(require(p, character.only = TRUE, quietly = TRUE)){
      message(paste("Pacote", p, "instalado e carregado com sucesso."))
    } else {
      message(paste("ERRO: Pacote", p, "nao pode ser carregado mesmo apos a instalacao."))
    }
  } else {
    message(paste("Pacote", p, "ja esta instalado. Carregando..."))
    if(require(p, character.only = TRUE, quietly = TRUE)){
      message(paste("Pacote", p, "carregado com sucesso."))
    } else {
      message(paste("ERRO: Pacote", p, "nao pode ser carregado."))
    }
  }
}


setwd(dir)
saida <- read_excel(xlsx, col_names = FALSE)

is_blank_row <- function(x) all(is.na(x) | trimws(as.character(x)) == "")

blocos <- list()
tempo <- 1

indices_end <- which(apply(saida, 1, function(x) any(grepl("~End", x, ignore.case = TRUE))))
inicio <- 1

for (i in seq_along(indices_end)) {
  bloco <- saida[inicio:(indices_end[i] - 1), , drop = FALSE]
  
  header_row <- which(apply(bloco, 1, function(x) any(grepl("Temperature", x, ignore.case = TRUE))))
  if (length(header_row) == 0) {
    inicio <- indices_end[i] + 1
    tempo <- tempo + 1
    next
  }
  header_row <- header_row[1]
  amostras <- as.character(bloco[header_row, -(1:2)])
  amostras <- amostras[!is.na(amostras) & amostras != ""]
  
  medidas <- bloco[(header_row + 1):nrow(bloco), , drop = FALSE]
  if (nrow(medidas) == 0) {
    inicio <- indices_end[i] + 1
    tempo <- tempo + 1
    next
  }
  
  blank_idx <- which(apply(medidas, 1, is_blank_row))
  split_indices <- if (length(blank_idx) == 0) nrow(medidas) + 1 else c(blank_idx, nrow(medidas) + 1)
  
  inicio_sub <- 1
  comprimento_vals <- comp_ondas_pre_def
  
  sub_count <- 1
  for (j in split_indices) {
    if ((j - 1) < inicio_sub) { inicio_sub <- j + 1; next }
    sub <- medidas[inicio_sub:(j - 1), , drop = FALSE]
    sub <- sub[!apply(sub, 1, is_blank_row), , drop = FALSE]
    
    if (nrow(sub) > 0 && sub_count <= length(comprimento_vals)) {
      sub <- sub[, 1:(length(amostras) + 2), drop = FALSE]
      dados_sub <- data.frame(
        Leitura = sub[[1]],
        Temperatura = sub[[2]],
        sub[, -c(1, 2), drop = FALSE]
      )
      colnames(dados_sub)[-c(1, 2)] <- amostras
      
      dados_long <- dados_sub %>%
        pivot_longer(cols = all_of(amostras), names_to = "Amostra", values_to = "Valor") %>%
        mutate(
          Tempo = tempo,
          Comprimento_onda = comprimento_vals[sub_count]
        )
      
      blocos[[length(blocos) + 1]] <- dados_long
      sub_count <- sub_count + 1
    }
    inicio_sub <- j + 1
  }
  
  tempo <- tempo + 1
  inicio <- indices_end[i] + 1
}

dados_final <- data.frame(bind_rows(blocos))
dados <- na.omit(dados_final)

# Ordena por amostra e tempo
dados <- dados[order(dados$Tempo, dados$Amostra), ]

# Converte colunas pra numerico
dados[] <- lapply(names(dados), function(x) if(x != "Amostra") as.numeric(dados[[x]]) else dados[[x]])

# media por amostra, dentro de tempo e de comprimento de onda
dados_summary <- dados %>%
  group_by(Tempo, Amostra, Comprimento_onda) %>%
  summarise(Media_Valor = mean(Valor), .groups = "drop")

final1 <- data.frame(dados_summary)
# ordenar amostras
final1$Amostra <- factor(final1$Amostra, levels = mixedsort(unique(final1$Amostra)))
final1 <- final1[order(final1$Tempo, final1$Amostra), ]
# por colunas
final2 <- data.frame(pivot_wider(final1, names_from = Amostra, values_from = Media_Valor))
final2 <- final2[, c("Comprimento_onda", setdiff(names(final2), "Comprimento_onda"))]

# por comprimento de onda (seguindo a pre-definicao)
lista_co <- lapply(comp_ondas_pre_def, function(x) subset(final2, Comprimento_onda == x))
names(lista_co) <- paste0("co", comp_ondas_pre_def)

# transposto
final3 <- final2[order(final2$Comprimento_onda), ]
final3t <- t(final3)


# Exportar planilhas
write.xlsx(final1, "resultados.xlsx")
write.xlsx(final2, "resultados_colunas.xlsx")
lapply(names(lista_co), function(n) {write.xlsx(lista_co[[n]], paste0("resultados_colunas_compr_onda_", gsub("co","",n), ".xlsx"))})
write.xlsx(final3t, "resultados_transposto.xlsx", colNames=FALSE, rowNames=TRUE)
# FIM

