library(tidyverse)
library(openxlsx)
library(officer)
library(stringr)
library(openxlsx)

# Tabla 
if(file.exists("Estructura fichas.xlsx")){
  
  tabla_fichas <- read.xlsx(
    "Estructura fichas.xlsx"
  )
  
} else {
  
  tabla_fichas <- tibble()
  
}

# Archivos
archivos <- list.files(
  "C:/Sofia/IJD/PROYECTO MIN-MAX/fichas",
  pattern="\\.docx$",
  full.names=TRUE
)

#Función para una ficha de word
extraer_ficha <- function(archivo){
  
  doc <- read_docx(archivo)
  
  texto <- docx_summary(doc)$text %>%
    paste(collapse="\n")
  
  
  # -------------------------
  # FUNCIÓN CAMPO
  # -------------------------
  extraer_campo <- function(nombre){
    
    patron <- paste0(
      nombre,
      "\\s*(\\(.*?\\))?\\s*:?\\s*(.*?)",
      "(?=\\n(?:Referencia bibliográfica|",
      "Ubicación en carpeta|",
      "Período cubierto por el trabajo.*?|",
      "Disciplinas involucradas|",
      "Tipo de trabajo|",
      "Palabras clave.*?|",
      "Resumen con foco.*?|",
      "Colaborador|",
      "Actores involucrados|",
      "Observaciones)|$)"
    )
    
    valor <- str_extract(
      texto,
      regex(patron, dotall = TRUE, ignore_case = TRUE)
    )
    
    valor <- str_remove(
      valor,
      regex(paste0("^", nombre, "\\s*(\\(.*?\\))?\\s*:"),
            ignore_case = TRUE)
    )
    
    trimws(valor)
  }
  
  # -------------------------
  # CAMPOS BASE (ORDEN CORRECTO)
  # -------------------------
  
  referencia <- extraer_campo("Referencia bibliográfica")
  tipo <- extraer_campo("Tipo de trabajo")
  actores <- extraer_campo("Actores involucrados")
  palabras <- extraer_campo("Palabras clave")
  resumen <- extraer_campo("Resumen con foco en temas de interés del proyecto")
  colaborador <- extraer_campo("Colaborador")
  
  # -------------------------
  # RESUMEN LIMPIO
  # -------------------------
  
  resumen <- str_remove(
    resumen,
    regex("^Resumen con foco[^\\n]*\\n?", ignore_case = TRUE)
  )
  
  resumen <- trimws(resumen)
  
  # -------------------------
  # AUTOR (ANTES DE USARLO)
  # -------------------------
  
  autor <- str_extract(referencia, "^[^,]+")
  
  # -------------------------
  # TÍTULO
  # -------------------------
  
  titulo <- str_extract(referencia, '["“](.*?)["”]')
  titulo <- str_remove_all(titulo, '["“”]')
  
  if(is.na(titulo) | titulo == ""){
    titulo <- tools::file_path_sans_ext(basename(archivo))
  }
  
  # -------------------------
  # FECHA
  # -------------------------
  
  fecha <- str_extract_all(referencia, "\\b(18|19|20)\\d{2}\\b")
  fecha <- tail(unlist(fecha), 1)
  
  # -------------------------
  # LIMPIEZA DE PALABRAS CLAVE
  # -------------------------
  
  palabras <- str_remove(
    palabras,
    regex("^Palabras clave\\s*\\(.*?\\)\\s*:", ignore_case = TRUE)
  )
  
  palabras <- trimws(palabras)
  
  # -------------------------
  # ETIQUETAS BASE (SOLO CONTROLADO)
  # -------------------------
  
  etiquetas <- c()
  
  # actores
  if(!is.na(actores) && actores != ""){
    
    actores_vec <- str_split(actores, ";|,")[[1]]
    actores_vec <- trimws(actores_vec)
    
    etiquetas <- c(etiquetas, actores_vec)
  }
  
  # palabras clave
  if(!is.na(palabras) && palabras != ""){
    
    palabras_vec <- str_split(palabras, ";|,")[[1]]
    palabras_vec <- trimws(palabras_vec)
    
    etiquetas <- c(etiquetas, palabras_vec)
  }
  
  # -------------------------
  # LIMPIEZA FINAL ROBUSTA
  # -------------------------
  
  etiquetas <- etiquetas[etiquetas != ""]
  etiquetas <- trimws(etiquetas)
  
  # quitar encabezados y basura
  etiquetas <- etiquetas[!grepl(
    "Referencia|Bibliográfica|Resumen|Observaciones|Colaborador|Período|Ubicación|Tipo|Trabajo",
    etiquetas,
    ignore.case = TRUE
  )]
  
  # quitar duplicados
  etiquetas <- unique(etiquetas)
  
  etiquetas <- sort(etiquetas)
  
  etiquetas <- paste(etiquetas, collapse = "; ")
  
  # -------------------------
  # OUTPUT
  # -------------------------
  
  tibble(
    Titulo = titulo,
    Autor = autor,
    Materia = palabras,
    Descripcion = resumen,
    Fecha = fecha,
    Tipo = tipo,
    Cobertura = "República Oriental del Uruguay",
    Editor = "Sala Uruguay de la Biblioteca Nacional",
    Fuente = NA,
    Colaborador = colaborador,
    Etiqueta = etiquetas
  )
}
# PROCESAR FICHAS
#########################################

nuevas_fichas <- purrr::map_dfr(
  
  archivos,
  
  extraer_ficha
  
)


#########################################
# UNIR SIN DUPLICAR
#########################################

tabla_fichas_final <- nuevas_fichas %>%
  
  distinct(
    
    Autor,
    Titulo,
    Fecha,
    
    .keep_all = TRUE
    
  )


#########################################
# GUARDAR
#########################################

write.xlsx(
  
  tabla_fichas_final,
  
  "Estructura fichas.xlsx",
  
  overwrite = TRUE
  
)



View(tabla_fichas_final)
write.xlsx(
  tabla_fichas_final,
  "tabla_fichas_final.xlsx",
  overwrite = TRUE
)