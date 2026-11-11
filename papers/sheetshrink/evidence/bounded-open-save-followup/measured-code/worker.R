args <- commandArgs(trailingOnly=TRUE)
library_name <- args[1]
input <- args[2]
state <- args[3]
output <- args[4]
suppressPackageStartupMessages(library(jsonlite))

publish <- function(phase, ...) {
  temporary <- paste0(state, ".pending")
  write_json(c(list(phase=phase), list(...)), temporary, auto_unbox=TRUE)
  file.copy(temporary, state, overwrite=TRUE)
  unlink(temporary)
}

phase <- "importing"
opened <- NULL
tryCatch({
  publish(phase)
  if (library_name == "r-openxlsx") {
    suppressPackageStartupMessages(library(openxlsx))
    phase <- "opening"
    publish(phase)
    started <- proc.time()[["elapsed"]]
    value <- loadWorkbook(input)
    opened <- proc.time()[["elapsed"]] - started
    phase <- "saving"
    publish(phase, open_seconds=opened)
    started <- proc.time()[["elapsed"]]
    saveWorkbook(value, output, overwrite=TRUE)
  } else {
    suppressPackageStartupMessages(library(readxl))
    suppressPackageStartupMessages(library(writexl))
    sheets <- excel_sheets(input)
    phase <- "opening"
    publish(phase)
    started <- proc.time()[["elapsed"]]
    value <- lapply(sheets, function(sheet) read_excel(input, sheet=sheet))
    names(value) <- sheets
    opened <- proc.time()[["elapsed"]] - started
    phase <- "saving"
    publish(phase, open_seconds=opened)
    started <- proc.time()[["elapsed"]]
    write_xlsx(value, output)
  }
  publish("completed", open_seconds=opened, save_seconds=proc.time()[["elapsed"]] - started)
}, error=function(error) {
  publish("failed", failed_phase=phase, open_seconds=opened, error_type="R error", error=conditionMessage(error))
  quit(status=1)
})
