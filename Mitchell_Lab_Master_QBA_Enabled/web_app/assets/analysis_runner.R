args <- commandArgs(trailingOnly=TRUE)
source(args[1], encoding="UTF-8")
options(warn=1)
need <- if (cfg$kind == "venn") c("readxl","dplyr","VennDiagram","openxlsx","tibble") else c("readxl","dplyr","ggplot2","stringr","ggrepel")
missing <- need[!vapply(need, requireNamespace, logical(1), quietly=TRUE)]
if (length(missing)) stop("Missing R packages: ", paste(missing, collapse=", "), ". Use Install R Packages on the analysis screen.")
for (pkg in need) suppressPackageStartupMessages(library(pkg, character.only=TRUE))
load_functions <- function(path) {
  for (expr in parse(path, encoding="UTF-8")) {
    if (is.call(expr) && identical(expr[[1]], as.name("<-")) &&
        length(expr)==3 && is.call(expr[[3]]) && identical(expr[[3]][[1]], as.name("function"))) eval(expr, envir=.GlobalEnv)
  }
}
selected_sheets <- function(path) {
  available <- readxl::excel_sheets(path)
  wanted <- unlist(cfg$sheets)
  if (!length(wanted)) return(available)
  absent <- setdiff(wanted, available)
  if (length(absent)) stop("Sheets not found in ", basename(path), ": ", paste(absent, collapse=", "))
  wanted
}
if (cfg$kind == "venn") {
  load_functions(file.path(cfg$assets, "venn_functions.R"))
  if (cfg$stage == "prepare") {
    sets <- list(); manifest <- list()
    for (path in unlist(cfg$files)) {
      all_sheets <- readxl::excel_sheets(path)
      for (sheet in selected_sheets(path)) {
        message("Reading ", basename(path), " / ", sheet)
        value <- read_set_from_excel_sheet(path, sheet)
        if (!length(value)) stop("Empty set: ", basename(path), " / ", sheet)
        name <- if (length(all_sheets)==1) tools::file_path_sans_ext(basename(path)) else sheet
        name <- make.unique(c(names(sets), name))[length(sets)+1]
        # Delimiter characters are reserved by the original intersection helper.
        name <- gsub("∩", "-", name, fixed=TRUE)
        if (grepl(" Only$", name)) name <- paste0(name, " set")
        name <- make.unique(c(names(sets), name))[length(sets)+1]
        sets[[name]] <- value
        manifest[[length(manifest)+1]] <- data.frame(name=name, count=length(value), file=basename(path), sheet=sheet)
      }
    }
    if (length(sets)<2) stop("At least two nonempty sets are needed; upload more workbooks or use multiple sheets.")
    saveRDS(sets, file.path(cfg$job, "sets.rds"))
    write.table(do.call(rbind,manifest), file.path(cfg$job,"sets.tsv"), sep="\t", quote=TRUE, row.names=FALSE, fileEncoding="UTF-8")
    message("READY: Select comparisons in the popup.")
  } else {
    sets <- readRDS(file.path(cfg$job, "sets.rds"))
    aliases <- unlist(cfg$aliases)
    if (length(aliases)!=length(sets) || anyDuplicated(aliases) || any(!nzchar(trimws(aliases)))) stop("Set names must be nonempty and unique.")
    if (any(grepl("∩", aliases, fixed=TRUE) | grepl(" Only$", aliases))) stop("Set names may not contain ∩ or end in ' Only'.")
    names(sets) <- aliases
    excel_dir <- file.path(cfg$job,cfg$excel_folder); plot_dir <- file.path(cfg$job,cfg$plot_folder)
    dir.create(excel_dir,recursive=TRUE,showWarnings=FALSE); dir.create(plot_dir,recursive=TRUE,showWarnings=FALSE)
    for (i in seq_along(cfg$comparisons)) {
      item <- cfg$comparisons[[i]]; idx <- as.integer(unlist(item$indices))
      if (length(idx)<2 || length(idx)>4 || anyDuplicated(idx) || any(!idx %in% seq_along(sets))) stop("Invalid comparison indices.")
      x <- sets[idx]
      message("Comparison ",i,"/",length(cfg$comparisons),": ", paste(names(x),collapse=" ∩ "))
      stub <- sprintf("%03d_%s_%s",i,cfg$prefix,item$name)
      excel <- file.path(excel_dir,paste0(stub,".xlsx"))
      plot <- file.path(plot_dir,stub)
      write_venn_excel(x,venn_group_table(x),excel)
      create_venn_plot_files(x,plot)
      expected <- c(excel,paste0(plot,c(".png",".pdf","_unlabeled.pdf")))
      if (!all(file.exists(expected)) || any(file.info(expected)$size<=0)) stop("Expected Venn output missing.")
    }
    message("COMPLETE: All Venn comparisons saved.")
  }
} else {
  load_functions(file.path(cfg$assets,"manhattan_functions.R"))
  results <- list(); successes <- 0L
  for (fi in seq_along(cfg$files)) {
    path <- cfg$files[[fi]]
    target <- file.path(cfg$job,sprintf("%03d_%s",fi,clean_file_stub(tools::file_path_sans_ext(basename(path)))))
    dir.create(target,recursive=TRUE,showWarnings=FALSE)
    all_sheets <- readxl::excel_sheets(path)
    for (si in seq_along(selected_sheets(path))) {
      sheet <- selected_sheets(path)[si]
      message("Workbook ",fi,"/",length(cfg$files),"; sheet ",sheet)
      name <- if (length(all_sheets)==1) tools::file_path_sans_ext(basename(path)) else sheet
      name <- paste0(cfg$prefix,"_",sprintf("%03d_",si),name)
      error <- NULL
      result <- tryCatch(make_manhattan_plot(path,sheet,name,target),error=function(e) {error <<- conditionMessage(e); NULL})
      if (!is.null(result)) {
        expected <- file.path(target,paste0(clean_file_stub(name),"_",cfg$suffix,c(".png",".tiff")))
        if (!all(file.exists(expected)) || any(file.info(expected)$size<=0)) {result<-NULL;error<-"Expected plot file missing"}
      }
      ok <- !is.null(result)
      if (ok) successes <- successes+1L
      reason <- if (ok) "Saved PNG and TIFF" else if (is.null(error)) "No usable rows / required columns missing" else error
      results[[length(results)+1]] <- data.frame(file=path,sheet=sheet,status=if(ok) "success" else "failed",detail=reason)
      if (!ok) message("FAILED SHEET: ", sheet, ": ",reason)
    }
  }
  table <- do.call(rbind,results)
  write.table(table,file.path(cfg$job,"sheet_results.tsv"),sep="\t",quote=TRUE,row.names=FALSE,fileEncoding="UTF-8")
  if (!successes) stop("No sheets produced plots. See sheet_results.tsv and run.log.")
  if (any(table$status!="success")) stop("PARTIAL RESULTS: Some sheets failed. Successful plots are retained. See sheet_results.tsv.")
  message("COMPLETE: All Manhattan sheets saved.")
}
