library(dplyr)
library(readxl)
library(purrr)
library(tidyr)
library(lubridate)

# Definindo a data inicial
inicial <- "2024-04-01"

# Definir caminho e listar arquivos
caminho <- "D:/Projetos/sanfelice/Dados"
arquivos <- list.files(path = caminho, pattern = "\\.xlsx$", full.names = TRUE)

# Carregar os arquivos em uma lista de data frames
lista_dfs <- map(arquivos, read_excel)

# Combinar todos os data frames em um único
dados_combinados <- bind_rows(lista_dfs)

# Filtrar e transformar os dados
dados <- dados_combinados %>% 
  filter(!is.na(Situação)) %>% 
  filter(Situação != "Cancelada") %>% 
  filter(Cliente == "NOVO BANCO CONTINENTAL S/A - BANCO MULTIPLO") %>% 
  mutate(Classificacao = ifelse(`Valor Bruto` == 8465.93, "CONSULTORIA", 
                                ifelse(`Valor Bruto` == 660.00, "CIVEL", "TRABALHISTA"))) %>% 
  filter(as.Date(Data) >= as.Date(inicial))

# Resumir os dados
dados_resumidos <- dados %>%
  mutate(MesAno = floor_date(Data, "month")) %>%  # Converte a data para mês e ano
  group_by(MesAno, Classificacao) %>%             # Agrupa por mês e classificação
  summarise(Soma_Valor_Bruto = sum(`Valor Bruto`, na.rm = TRUE)) %>% # Resumir o valor bruto
  ungroup()

# Transformar de formato longo para largo
dados_largos <- dados_resumidos %>%
  pivot_wider(names_from = Classificacao, values_from = Soma_Valor_Bruto) %>%
  replace_na(list(CONSULTORIA = 0, CIVEL = 0, TRABALHISTA = 0))

# Somatórios por coluna
classificacoes <- c("CONSULTORIA", "CIVEL", "TRABALHISTA")
totais_bruto <- colSums(dados_largos[, -1], na.rm = TRUE)
totais_bruto[setdiff(classificacoes, names(totais_bruto))] <- 0  # Garantir todas as classificações

# Cálculos tributários
carga_tributaria <- 0.156 * totais_bruto
total_liquido <- totais_bruto - carga_tributaria
percentual_participacao <- 0.10
valor_participacao <- percentual_participacao * total_liquido
total_participacao <- sum(valor_participacao, na.rm = TRUE)

# Montagem da tabela final
tabela <- data.frame(
  Categoria = c("TOTAL BRUTO", "CARGA TRIBUTÁRIA 15,60%", "TOTAL LIQUIDO", 
                "PERCENTUAL DE PARTICIPAÇÃO", "VALOR DE PARTICIPAÇÃO"),
  CIVEL = c(totais_bruto["CIVEL"], carga_tributaria["CIVEL"], total_liquido["CIVEL"], 
            percentual_participacao * 100, valor_participacao["CIVEL"]),
  CONSULTORIA = c(totais_bruto["CONSULTORIA"], carga_tributaria["CONSULTORIA"], 
                  total_liquido["CONSULTORIA"], percentual_participacao * 100, 
                  valor_participacao["CONSULTORIA"]),
  TRABALHISTA = c(totais_bruto["TRABALHISTA"], carga_tributaria["TRABALHISTA"], 
                  total_liquido["TRABALHISTA"], percentual_participacao * 100, 
                  valor_participacao["TRABALHISTA"])
)

# Visualizar a tabela
print(tabela)



# Valores já calculados
total_participacao <- sum(valor_participacao, na.rm = TRUE)

# Cálculo dos impostos
irrf <- total_participacao * 0.015  # 1,5% de IRRF
csll_pis_cofins <- total_participacao * 0.0465  # 4,65% de CSLL/PIS/COFINS
total_liquido_final <- total_participacao - irrf - csll_pis_cofins

# Montar o data frame final
resumo_final <- data.frame(
  Categoria = c("TOTAL DE PARTICIPAÇÃO", "", "VALOR PARA EMISSÃO DE NOTA FISCAL", 
                "VALOR BRUTO", "IRRF (1,5%)", "CSLL/PIS/COFINS (4,65%)", "TOTAL LIQUIDO"),
  Valor = c(format(total_participacao, nsmall = 2, big.mark = ","), 
            "",  # Linha em branco para espaçamento
            "",  # Linha em branco para "VALOR PARA EMISSÃO DE NOTA FISCAL"
            format(total_participacao, nsmall = 2, big.mark = ","), 
            format(irrf, nsmall = 2, big.mark = ","), 
            format(csll_pis_cofins, nsmall = 2, big.mark = ","), 
            format(total_liquido_final, nsmall = 2, big.mark = ","))
)

# Visualizar o resumo final
print(resumo_final, row.names = FALSE)



#############

setwd("D:/Projetos/sanfelice")
library(openxlsx)

# Criar um novo arquivo Excel
wb <- createWorkbook()

# Adicionar uma única planilha
addWorksheet(wb, "Resumo Completo")

# Escrever 'dados_largos' na planilha, começando na célula A1
writeData(wb, sheet = "Resumo Completo", dados_largos, startRow = 1, startCol = 1)

# Calcular o número de linhas ocupadas por 'dados_largos' para determinar onde escrever o próximo bloco de dados
linhas_dados_largos <- nrow(dados_largos) + 2  # +2 para espaçamento

# Escrever 'tabela' na planilha, começando logo abaixo de 'dados_largos'
writeData(wb, sheet = "Resumo Completo", tabela, startRow = linhas_dados_largos, startCol = 1)

# Calcular o número de linhas ocupadas por 'tabela'
linhas_tabela <- linhas_dados_largos + nrow(tabela) + 2  # +2 para espaçamento

# Escrever 'resumo_final' na planilha, começando logo abaixo de 'tabela'
writeData(wb, sheet = "Resumo Completo", resumo_final, startRow = linhas_tabela, startCol = 1)

# Salvar o arquivo Excel
saveWorkbook(wb, file = "Resumo_Final_Unico.xlsx", overwrite = TRUE)
