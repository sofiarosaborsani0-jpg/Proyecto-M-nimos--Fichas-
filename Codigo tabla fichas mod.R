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
  
  
  extraer_campo <- function(nombre){
    
    patron <- paste0(
      nombre,
      "\\s*(\\(.*?\\))?\\s*:?\\s*(.*?)",
      "(?=\\n(?:Referencia bibliográfica|",
      "Ubicación en carpeta|",
      "Período cubierto.*?|",
      "Disciplinas involucradas|",
      "Tipo de trabajo|",
      "Palabras clave.*?|",
      "Resumen con foco.*?|",
      "Actores involucrados|",
      "Colaborador|",
      "Observaciones)|$)"
    )
    
    valor <- str_extract(
      texto,
      regex(
        patron,
        dotall=TRUE,
        ignore_case=TRUE
      )
    )
    
    valor <- str_remove(
      
      valor,
      
      regex(
        
        paste0(
          "^",
          nombre,
          "\\s*(\\(.*?\\))?",
          "\\s*:?"      # ← ":" opcional
        ),
        
        ignore_case = TRUE
        
      )
      
    )
    
    trimws(valor)
    
  }
  
  
  #############
  # CAMPOS
  #############
  
  referencia <- extraer_campo("Referencia bibliográfica")
  
  palabras <- extraer_campo("Palabras clave")
  
  resumen <- extraer_campo(
    "Resumen con foco en temas de interés del proyecto"
  )
  
  actores <- extraer_campo(
    "Actores involucrados"
  )
  
  tipo <- extraer_campo("Tipo de trabajo")
  
  colaborador <- extraer_campo("Colaborador")
  
  
  #############
  # LIMPIAR MATERIA
  #############
  
  palabras <- str_remove(
    
    palabras,
    
    regex(
      
      "^Palabras\\s+clave\\s*(\\(.*?\\))?\\s*:?",
      
      ignore_case = TRUE
      
    )
    
  )
  
  palabras <- trimws(palabras)
  
  
  
  #############
  # TITULO
  #############
  
  titulo <- str_extract(
    referencia,
    '["“](.*?)["”]'
  )
  
  titulo <- str_remove_all(
    titulo,
    '["“”]'
  )
  
  if(is.na(titulo)){
    
    titulo <- tools::file_path_sans_ext(
      basename(archivo)
    )
    
  }
  
  
  #############
  # AUTOR
  #############
  
  autor <- str_extract(
    referencia,
    "^[^,]+"
  )
  
  
  
  #############
  # FECHA
  #############
  
  fecha <- str_extract_all(
    referencia,
    "\\b(18|19|20)\\d{2}\\b"
  )
  
  fecha <- tail(
    unlist(fecha),
    1
  )
  
  
  
  #############
  # LIMPIAR RESUMEN
  #############
  
  resumen <- str_remove(
    
    resumen,
    
    regex(
      
      "^Resumen\\s+con\\s+foco.*?\\)?\\s*:?",
      
      ignore_case = TRUE
      
    )
    
  )
  
  # eliminar puntos o saltos sobrantes al inicio
  resumen <- str_remove(
    resumen,
    "^[\\s\\n\\r\\.]+"
  )
  
  resumen <- trimws(resumen)
  
  
  
  #############
  # ETIQUETAS:
  # SOLO RESUMEN + ACTORES
  #############
  
  texto_etiquetas <- resumen
  
  
  # nombres propios
  etiquetas <- str_extract_all(
    
    texto_etiquetas,
    
    "(?:[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+(?:\\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+)+)"
    
  )
  
  etiquetas <- unlist(etiquetas)
  
  
  #############
  # AGREGAR ACTORES
  #############
  
  if(!is.na(actores)){
    
    actores_vec <- str_split(
      actores,
      ";|,"
    )[[1]]
    
    etiquetas <- c(
      etiquetas,
      trimws(actores_vec)
    )
    
  }
  
  
  #############
  # LIMPIEZA
  #############
  
  basura <- c(
    
    autor,
    titulo,
    
    "Palabras clave",
    "Resumen",
    "Referencia",
    
    "El","La","Los","Las",
    "En","Por","Sobre",
    
    "Trabajo",
    "Tipo"
    
  )
  
  etiquetas <- trimws(etiquetas)
  
  # sacar artículos iniciales
  etiquetas <- str_remove(
    etiquetas,
    regex(
      "^(El|La|Los|Las)\\s+",
      ignore_case = TRUE
    )
  )
  
  # quitar espacios dobles
  etiquetas <- str_squish(etiquetas)
  
  etiquetas <- etiquetas[
    !etiquetas %in% basura
  ]
  
  # eliminar duplicados después de limpiar
  etiquetas <- unique(etiquetas)
  
  etiquetas <- sort(etiquetas)
  
  etiquetas <- paste(
    etiquetas,
    collapse="; "
  )
  
  
  
  #############
  # OUTPUT
  #############
  
  tibble(
    
    Titulo=titulo,
    
    Autor=autor,
    
    Materia=palabras,
    
    Descripcion=resumen,
    
    Fecha=fecha,
    
    Tipo=tipo,
    
    Cobertura=
      "República Oriental del Uruguay",
    
    Editor=
      "Sala Uruguay de la Biblioteca Nacional",
    
    Fuente=NA,
    
    Colaborador=colaborador,
    
    Etiqueta=etiquetas
    
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