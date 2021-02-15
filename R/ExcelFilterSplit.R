#' ExcelFilterSplit
#'
#' This package will help you create multiple filtered Excel/CSV files. Filtering is made by choosing a column ; one output file per unique value of the selected column (not case sensitive ; not sensitive to accents).
#' If you choose Excel output format, you have different styles available, but you can also keep the format and style of your original Excel input file.
#' Output files are compressed into a Zip folder. The file will be export into your current working directory.
#'
#' @details
#' In your Excel sheet, make sure the data starts in cell A2 and columnns in cell A1.
#' The output zip file containing your multiple output files will be created in your current working directory.
#' Same CSV separators are used for input and output CSV files
#'
#' @param input_path a character string (file path) that ends in ".xlsx" or ".xls" or ".csv". Note that the file must be local (a URL does not work).
#' @param input_filterby a character string (column name) or a numeric integrer (column index)
#' @param output_names1 a character string, or NULL (for no value). Output names of type: **names1 ( filterby ) names2 .extension**
#' @param output_names2 a character string, or NULL (for basename(input_path) without ext.)
#' @param output_format a character string : "default" (for the input file format) or "excel" or "csv"
#' @param output_cols a list of character string (columns names) or a list of numeric integrers (columns indexes) or NULL (for all columns)
#' @param CSVsep a list of 2 character string (field sep , decimal sep). Ignored if not CSV input file and/or CSV output files.
#' @param input_EXCsheet a character string (sheet name) or a numeric vector (sheet index). Ignored if CSV input file.
#' @param output_EXCstyle a character string : one of Excel TableStyles or "default" (for "keep_style" if Excel input; for "TableStyleMedium18" if CSV input ) or "none" or "keep_style". Ignored if CSV output files.
#' @param output_EXCsheets a list of character string (sheet names) or a list of numeric integrers (sheet indexes) or NULL (for no sheet). Ignored if CSV output files.
#' @param output_maxfiles a numeric vector
#' @param ShinyApp_zipfilename a character string (file path) that ends in ".zip". Needed for a use in a ShinyApp environment with a DownloadHandler. It will also activate the shiny::incProgress in case you want to add a progress bar in your app.
#'
# @return None
#'
#'
#' @examples
#'
#' ## an Excel file for example
#'
#' data <- list(
#'   c(Company = "AA", EmployeeID = "1", Info = "Info1"),
#'   c(Company = "BB", EmployeeID = "2", Info = "Info2"),
#'   c(Company = "BB", EmployeeID = "3", Info = "Info3"),
#'   c(Company = "AA", EmployeeID = "4", Info = "Info4"),
#'   c(Company = "CC", EmployeeID = "5", Info = "Info5")
#' )
#' data <- t(as.data.frame(data))
#' colnames(data) <- c("Company", "EmployeeID", "Info")
#' rownames(data) <- NULL
#' head(data)
#'
#' getwd()
#' openxlsx::write.xlsx(data, file = "example.xlsx")
#'
#'
#' ## ExcelFilterSplit function
#'
#' MY_input_path <- "example.xlsx"
#' MY_input_filterby <- "Company"
#'
#'
#' ExcelFilterSplit::ExcelFilterSplit(
#'   input_path = MY_input_path,
#'   input_filterby = MY_input_filterby
#' )
#'
#'
# @note
#' @author Thibault Maurin (\email{tmaurin.pro@@gmail.com})
#'
# @references \url{http://en.wikipedia.org/wiki/}
#'
#' @seealso \code{\link{ExcelFilterSplit}}
#'
#' @keywords ExcelFilerSplit
#'
#'
#' @import readxl
#' @import openxlsx
#' @import tools
#' @import stringi
#' @import itertools2
#' @import fs
#' @importFrom shiny incProgress
#' @importFrom stats na.omit
#' @importFrom utils head read.csv write.table
#'
#' @export
#'
#'
ExcelFilterSplit <- function(

                             ## Arguments -----------------------------------------

                             input_path,
                             input_filterby,
                             output_names1 = NULL,
                             output_names2 = NULL,
                             output_format = "excel",
                             output_cols = NULL,
                             CSVsep = c(",", "."),
                             input_EXCsheet = 1,
                             output_EXCstyle = "default",
                             output_EXCsheets = NULL,
                             output_maxfiles = 269,
                             ShinyApp_zipfilename = NULL

) {

  print((" "))
  print("You have started using ExcelFilterSplit function !")
  print((" "))
  print("-- LOADING ExcelFilterSplit...")
  print("-- STARTING VERIFICATIONS (1/2)... ")



  ## Variables & functions  --------------------------------------------

  if (!is.null(ShinyApp_zipfilename)) {
    shiny::incProgress(
      amount = 1 / 12,
      message = "Doing verifications..."
    )
  }

  assert_error <- function(expression, message) {
    if (missing(message)) {
      message <- paste0("Condition ", deparse(as.list(match.call())$expression), " is not TRUE")
    }
    if (!expression) {
      stop(message, call. = FALSE)
    }
  }


  SheetExcel_StyleObjects <- function(
                                      nameExcelSheet,
                                      styleObjects) {
    styles_sheet_to_modify <- c()

    for (i in 1:length(styleObjects)) {
      if (nameExcelSheet %in% styleObjects[[i]]$sheet) {
        styles_sheet_to_modify <- c(styles_sheet_to_modify, i)
      }
    }
    return_value <- styleObjects[styles_sheet_to_modify]
    return(return_value)
  }



  # CHECKS : ---------------------------------------------------------------------------------------------

  ## input_path ----

  assert_error(
    expression = is.character(input_path),
    message = "Invalid argument 'input_path' (path of the file).
              Argument format accepted : string "
  )
  assert_error(
    expression = endsWith(input_path, c(".csv")) | endsWith(input_path, c(".xlsx")) | endsWith(input_path, c(".xls")),
    message = "Invalid argument 'input_path' (path of the file).
              File formats accepted : .csv .xlsx .xls"
  )


  if (endsWith(input_path, ".csv")) {
    truevalue_inputformat <- "csv"
  } else {
    truevalue_inputformat <- "excel"
  }

  print(paste0("Checking input_path", " : ", input_path))
  #print(paste0("truevalue_inputformat", " : ", truevalue_inputformat))



  ## output_format ----

  assert_error(
    expression = isTRUE(output_format == "default") | isTRUE(output_format == "csv") | isTRUE(output_format == "excel"),
    message = "Invalid argument 'output_format' (format files created).
              Argument format accepted : 'default' (same format as your input file) or 'csv' or 'excel'."
  )

  if (output_format == "default" && truevalue_inputformat == "excel") {
    truevalue_outputformat <- "excel"
  } else if (output_format == "default" && truevalue_inputformat == "csv") {
    truevalue_outputformat <- "csv"
  } else {
    truevalue_outputformat <- output_format
  }

  print(paste0("Checking output_format", " : ", output_format))
  #print(paste0("truevalue_outputformat", " : ", truevalue_outputformat))



  ## output_EXCstyle ----

  defaultvalue_ExcelTableStyle <- "TableStyleMedium18"
  values_ExcelTableStyleNames <- c()
  for (i in c(1:21)) {
    values_ExcelTableStyleNames <- c(values_ExcelTableStyleNames, paste0("TableStyleLight", as.character(i)))
  }
  for (i in c(1:28)) {
    values_ExcelTableStyleNames <- c(values_ExcelTableStyleNames, paste0("TableStyleMedium", as.character(i)))
  }
  for (i in c(1:11)) {
    values_ExcelTableStyleNames <- c(values_ExcelTableStyleNames, paste0("TableStyleDark", as.character(i)))
  }
  # https://github.com/xuri/excelize-doc/blob/master/en/utils.md

  print(paste0("output_EXCstyle", " : ", output_EXCstyle))

  assert_error(
    expression = isTRUE(output_EXCstyle == "default") | isTRUE(output_EXCstyle == "") | isTRUE(output_EXCstyle == "keep_style") | isTRUE(output_EXCstyle == "none") | isTRUE(output_EXCstyle %in% values_ExcelTableStyleNames),
    message = "Invalid argument 'output_EXCstyle' (style of Excel files created).
              Argument format accepted : 'default' or 'none' or 'keep_style' or one of Excel TableStyles.
              See link https://github.com/xuri/excelize-doc/blob/master/en/utils.md for Excel TableStyles availabilities."
  )
  print(paste0("Checking output_EXCstyle", " : ", output_EXCstyle))



  ## input_filterby (1/2) -----

  assert_error(
    expression = is.character(input_filterby) | is.numeric(input_filterby) && input_filterby > 0,
    message = "Invalid argument 'input_filterby' (column feature use to filter file/sheet).
              Argument format accepted : string or integrer >0."
  )
  print(paste0("Checking input_filterby", " : ", input_filterby))



  ## output_cols (1/2) -----

  if (!is.null(output_cols)) {
    assert_error(
      expression = is.character(output_cols) | is.numeric(output_cols) && output_cols > 0,
      message = "Invalid argument 'output_cols' (columns to keep in the output files).
              Argument format accepted : vector of strings(Column names) or integrers >0 (Column indexes),
      or 'NULL' for all columns"
    )
    print(paste(c("Checking output_cols: ", output_cols), collapse = " "))
  }


  ## CSVsep ----

  assert_error(
    expression = is.character(CSVsep) && length(CSVsep) == 2,
    message = "Invalid argument 'CSVsep'.
              Argument format accepted : list of 2 strings ( field separator character(s) ,
    decimal separator character(s) )."
  )

  print(paste0("Checking CSVsep[1]", " : ", CSVsep[1]))
  print(paste0("Checking CSVsep[2]", " : ", CSVsep[2]))


  ## input_EXCsheet, output_EXCsheets  ----

  if (truevalue_inputformat == "excel") {
    listsheetsfile <- readxl::excel_sheets(input_path)

    assert_error(
      expression = isTRUE(input_EXCsheet %in% listsheetsfile) | isTRUE(input_EXCsheet %in% 1:length(listsheetsfile)),
      message = paste(c("Invalid argument 'input_EXCsheet' : sheet out of range.
                    Possibles values : ", listsheetsfile), collapse = " ")
    )

    if (is.numeric(input_EXCsheet)) {
      input_EXCsheet <- listsheetsfile[input_EXCsheet]
    }

    if (!is.null(output_EXCsheets)) {
      assert_error(
        expression = all(output_EXCsheets %in% listsheetsfile) | all(output_EXCsheets %in% 1:length(listsheetsfile)),
        message = paste(c("Invalid argument 'output_EXCsheets' : sheet out of range.
                    Possibles values : ", listsheetsfile), collapse = " ")
      )

      if (isTRUE(is.numeric(output_EXCsheets))) {
        output_EXCsheets <- listsheetsfile[output_EXCsheets]
      }
    }
  }
  print(paste0("Checking input_EXCsheet", " : ", input_EXCsheet))
  print(paste(c("Checking output_EXCsheets", " : ", output_EXCsheets), collapse = " "))



  print("Verifications (1/2) completed ")

  # PART 1 : ---------------------------

  ## : load data ----

  print("-- LOADING INPUT FILE...")

  if (!is.null(ShinyApp_zipfilename)) {
    shiny::incProgress(
      amount = 1 / 12,
      message = "Reading data..."
    )
  }

  if (truevalue_inputformat == "excel") {
    if (truevalue_outputformat == "excel" && ( output_EXCstyle == "default" | output_EXCstyle == "")) {
      output_EXCstyle <- "keep_style"
    }
    if (truevalue_outputformat == "csv") {
      output_EXCstyle <- NULL
    }


    MY_file <- readxl::read_excel(
      path = input_path,
      sheet = input_EXCsheet,
      col_names = TRUE,
      col_types = NULL, # = guessed / or "text" ???
      na = c("", "#VALEUR!", "#VALUE!", "#NA"),
      trim_ws = TRUE,
      skip = 0
    )
  } else if (truevalue_inputformat == "csv") {
    input_EXCsheet <- NULL
    output_EXCsheets <- NULL
    if (truevalue_outputformat == "excel" && ( output_EXCstyle == "default" | output_EXCstyle == "")) {
      output_EXCstyle <- defaultvalue_ExcelTableStyle
    }
    if (truevalue_outputformat == "excel" && output_EXCstyle == "keep_style") {
      output_EXCstyle <- defaultvalue_ExcelTableStyle
    }
    if (truevalue_outputformat == "csv") {
      output_EXCstyle <- NULL
    }

    #+ try
    MY_file <- read.csv(
      file = input_path,
      sep = CSVsep[1],
      dec = CSVsep[2]
    )
  }


  MY_file <- as.data.frame(MY_file)
  colnames(MY_file) <- gsub("^[...]..", "Column_", colnames(MY_file))
  colnames(MY_file) <- gsub("[...]..", "_", colnames(MY_file))
  #print(head(MY_file))

  print("File loaded")

  print("-- STARTING VERIFICATIONS (2/2)... ")

  ## input_filterby , output_cols (2/2) ----

  #print(paste0("Checking input_filterby : ", input_filterby))
  #print(paste(c("Checking output_cols : ", output_cols), collapse = " "))


  colNames_MY_file <- names(MY_file)
  colIndex_MY_file <- c(1:length(colNames_MY_file))

  assert_error(
    expression = isTRUE(input_filterby %in% colNames_MY_file) | isTRUE(input_filterby %in% colIndex_MY_file),
    message = paste(c("Invalid argument 'input_filterby' : column not found.
            Possibles values : ", colNames_MY_file), collapse = " ")
  )


  if (is.null(output_cols)) {
    output_cols <- colIndex_MY_file
  } else {
    assert_error(
      expression = all(output_cols %in% colNames_MY_file) | all(output_cols %in% colIndex_MY_file),
      message = paste(c("Invalid argument 'output_cols' :  columns out of range.
            Possibles values : ", colIndex_MY_file), collapse = " ")
    )
  }


  if (isTRUE(is.numeric(output_cols))) {
    index_columns_to_keep <- sort(output_cols)
  } else if (isTRUE(is.character(output_cols))) {
    index_columns_to_keep <- match(output_cols, colNames_MY_file)
  }
  index_columns_to_keep <- unique(index_columns_to_keep)


  ## output_maxfiles  (2/2)  -----

  assert_error(
    expression = is.numeric(output_maxfiles) && output_maxfiles > 0,
    message = "Invalid argument 'output_maxfiles' (maximum number of files allowed) .
              Argument format accepted : integrer >0."
  )
  print(paste0("Checking output_maxfiles : ", output_maxfiles))

  nb_files_created <- length(unique(stringi::stri_trans_general(tolower(MY_file[, input_filterby]), "Latin-ASCII")))
  #print(paste0("Checking nb_files_created : ", nb_files_created))
  assert_error(
    expression = nb_files_created <= output_maxfiles,
    message = paste0("Too many files are going to be created; modify 'output_maxfiles' value to allowed more output files.
            You need ", nb_files_created, " files. Your max is currently set at : ", output_maxfiles)
  )


  ## truevalue_zipfileame ----

  assert_error(
    expression = is.character(output_names2) && length(output_names2) < 69 | is.null(output_names2),
    message = "Invalid argument 'output_names2'.
              Argument format accepted : string  <69 characters"
  )

  assert_error(
    expression = is.character(output_names1) && length(output_names1) < 69 | is.null(output_names1),
    message = "Invalid argument 'output_names1'.
              Argument format accepted : string  <69 characters"
  )


  if (is.null(output_names1)) {
    output_names1 <- ""
  } else {
    output_names1 <- paste0(output_names1, " ")
  }
  print(paste0("Checking output_names1", " : ", output_names1))

  if (is.null(output_names2)) {
    output_names2 <- sub(pattern = "(.*)\\..*$", replacement = "\\1", basename(input_path))
  }
  print(paste0("Checking output_names2", " : ", output_names2))

  truevalue_zipfileame <- paste0(output_names1, "(", input_filterby, ") ", output_names2)
  truevalue_zipfileame <- fs::path_sanitize(truevalue_zipfileame, replacement = "-")

  #print(paste0("truevalue_zipfileame", " : ", truevalue_zipfileame))

  print("Verifications (2/2) completed ")


  ## : create temp folder ----

  temp_folder_path <- paste0(
    tempdir(), "/",
    do.call(paste0, Map(stri_rand_strings, n = 1, length = c(8), pattern = c("[A-Z]")))
  )
  dir.create(temp_folder_path)
  #print(paste0("temp_folder_path", " : ", temp_folder_path))


  # PART 2 : ----------------------------------------------------------------------------------

  print("-- CREATING OUTPUT FILES...")

  if (!is.null(ShinyApp_zipfilename)) {
    shiny::incProgress(
      amount = 1 / 12, # 1/12,
      message = paste0("Creating files...", "\n\n")
    )
  }

  #print(dim(MY_file))
  #print(unique(stringi::stri_trans_general(tolower(MY_file[, input_filterby]), "Latin-ASCII")))

  nb_files_created <- length(unique(stringi::stri_trans_general(tolower(MY_file[, input_filterby]), "Latin-ASCII")))

  counter <- 0

  for (filter_by in unique(stringi::stri_trans_general(tolower(MY_file[, input_filterby]), "Latin-ASCII"))) {

    counter <- counter +1
    print(paste0("Creating file (", counter, "/", nb_files_created, ") : ", filter_by))


    filter_rows <- c((filter_by == stringi::stri_trans_general(tolower(MY_file[, input_filterby]), "Latin-ASCII")) %in% TRUE)

    filter_by_name <- unique(MY_file[filter_rows, input_filterby])[1]
    #print(paste0("filter_by_name : ", filter_by_name))

    ## : Create Excel files ----

    if (truevalue_outputformat == "excel") {
      filenaaame <- paste0(output_names1, "(", filter_by_name, ") ", output_names2, ".xlsx")
      filenaaame <- fs::path_sanitize(filenaaame, replacement = "-")
      #print(paste0("filename : ", filenaaame))

      if (!is.null(ShinyApp_zipfilename)) {
        shiny::incProgress(
          amount = 8.5 / nb_files_created / 12,
          detail = paste0("\n\n", filenaaame)
        )
      }

      if (truevalue_inputformat == "csv") {
        name_EXCsheet <- "Sheet1"
      } else if (truevalue_inputformat == "excel") {
        name_EXCsheet <- input_EXCsheet
      }

      MY_wb_copy <- openxlsx::createWorkbook() #("ExcelFilterSplit")
      openxlsx::addWorksheet(
        wb = MY_wb_copy,
        sheetName = name_EXCsheet)

      if (truevalue_inputformat == "excel") {
        indexshe <- match(name_EXCsheet, listsheetsfile)
      }

      if (truevalue_inputformat == "excel" && (output_EXCstyle == "keep_style" | output_EXCstyle == "none")) {
        wb4Style <- openxlsx::loadWorkbook(file = input_path)
        if (isTRUE(length(wb4Style$styleObjects) > 0)) {
          Style_exist <- TRUE
        } else {
          Style_exist <- FALSE
        }
      } else {
        Style_exist <- FALSE
      }


      ## :: Create WB filtered_data -----


      ## WB:: keep_style ----

      if (output_EXCstyle == "keep_style" && Style_exist) {
        Xdata <- MY_file[filter_rows, index_columns_to_keep]
        colnames(Xdata) <- gsub("^(Column_).*", "", colnames(Xdata))

        openxlsx::writeData(
          wb = MY_wb_copy,
          sheet = name_EXCsheet,
          x = Xdata
        )

        stylesObject_sheet_to_modify <- SheetExcel_StyleObjects(name_EXCsheet, wb4Style$styleObjects)
        stylesObject_new <- stylesObject_sheet_to_modify
        for (i in 1:length(stylesObject_sheet_to_modify)) {
          keeped_index <- (stylesObject_sheet_to_modify[[i]]$cols %in% index_columns_to_keep)
          stylesObject_new[[i]]$cols <- stylesObject_sheet_to_modify[[i]]$cols[keeped_index]
          stylesObject_new[[i]]$rows <- stylesObject_sheet_to_modify[[i]]$rows[keeped_index]
        }

        stylesObject_new2 <- stylesObject_new
        for (ii in 1:length(stylesObject_new)) {
          for (xxx in as.list(itertools2::izip(value = index_columns_to_keep, index = seq_along(index_columns_to_keep)))) {
            replace_index <- (stylesObject_new[[ii]]$cols == xxx$value)
            stylesObject_new2[[ii]]$cols[replace_index] <- xxx$index
          }
        }

        for (i in 1:length(stylesObject_new2)) {
          openxlsx::addStyle(
            wb = MY_wb_copy,
            stack = TRUE,
            sheet = stylesObject_new2[[i]]$sheet,
            style = stylesObject_new2[[i]]$style,
            rows = stylesObject_new2[[i]]$rows,
            cols = stylesObject_new2[[i]]$cols
          )
        }

        indexstyle <- wb4Style$colWidths[[indexshe]]

        if (length(indexstyle) != 0) {
          keeped_indexx <- (as.numeric(names(indexstyle)) %in% index_columns_to_keep)
          indexstyle <- indexstyle[keeped_indexx]
          names(indexstyle) <- c(1:length(index_columns_to_keep))

          openxlsx::setColWidths(MY_wb_copy,
            sheet = name_EXCsheet,
            cols = as.numeric(names(indexstyle)),
            widths = as.numeric(indexstyle)
          )
        }

        if (length(indexstyle) == 0) {
          openxlsx::setColWidths(MY_wb_copy,
            sheet = name_EXCsheet,
            cols = c(1:length(index_columns_to_keep)),
            widths = "auto"
          )
        }
      }

      ## WB:: others Styles ----

      if (isTRUE(output_EXCstyle %in% values_ExcelTableStyleNames)) {
        filtered_dataaa <- MY_file[filter_rows, index_columns_to_keep]
        colnames(filtered_dataaa) <- gsub("\\n", " ", colnames(filtered_dataaa))
        openxlsx::writeDataTable(
          wb = MY_wb_copy,
          sheet = name_EXCsheet,
          x = filtered_dataaa,
          tableStyle = output_EXCstyle
        )
        openxlsx::freezePane(
          wb = MY_wb_copy,
          sheet = name_EXCsheet,
          firstRow = TRUE
        )
        setColWidths(
          wb = MY_wb_copy,
          sheet = name_EXCsheet,
          cols = c(1:length(colnames(filtered_dataaa))),
          widths = "auto"
        )
      }

      if (output_EXCstyle == "none" | (output_EXCstyle == "keep_style" && !Style_exist)) {
        openxlsx::writeData(
          wb = MY_wb_copy,
          sheet = name_EXCsheet,
          x = MY_file[filter_rows, index_columns_to_keep],
          colNames = TRUE
        )
      }



      if (!is.null(output_EXCsheets)) {

        ## output_EXCsheets:: Keep_style ----

        if (output_EXCstyle == "keep_style" && Style_exist) {
          for (i in output_EXCsheets) {
            df_i <- readxl::read_xlsx(
              path = input_path,
              sheet = i,
              col_types = "text",
              col_names = FALSE,
              trim_ws = FALSE)
            dfirange <- paste0("A1:", openxlsx::getCellRefs(data.frame(nrow(df_i) + 6, ncol(df_i))))
            df_i <- readxl::read_xlsx(
              path = input_path,
              sheet = i,
              col_names = TRUE,
              trim_ws = TRUE,
              range = dfirange,
              col_types = NULL
            )

            colnames(df_i) <- gsub("[...].*", "", colnames(df_i))

            openxlsx::addWorksheet(
              wb = MY_wb_copy,
              sheetName = i
            )
            openxlsx::writeData(
              wb = MY_wb_copy,
              sheet = i,
              x = df_i,
              colNames = TRUE
            )
            #- rm(df_i)

            stylesObject_sheets_to_modify <- SheetExcel_StyleObjects(i, wb4Style$styleObjects)
            for (ii in 1:length(stylesObject_sheets_to_modify)) {
              openxlsx::addStyle(
                wb = MY_wb_copy,
                stack = TRUE,
                sheet = stylesObject_sheets_to_modify[[ii]]$sheet,
                style = stylesObject_sheets_to_modify[[ii]]$style,
                rows = stylesObject_sheets_to_modify[[ii]]$rows,
                cols = stylesObject_sheets_to_modify[[ii]]$cols
              )
            }

            index <- match(i, listsheetsfile)
            openxlsx::setColWidths(
              wb = MY_wb_copy,
              sheet = i,
              cols = ifelse(length(wb4Style$colWidths[[index]]) == 0,
                c(1:length(index_columns_to_keep)),
                as.numeric(names(wb4Style$colWidths[[index]]))
              ),
              widths = ifelse(length(wb4Style$colWidths[[index]]) == 0,
                "auto",
                as.numeric(wb4Style$colWidths[[index]])
              )
            )

            openxlsx::setRowHeights(
              wb = MY_wb_copy,
              sheet = i,
              rows = ifelse(length(wb4Style$rowHeights[[index]]) == 0,
                c(1:length(index_columns_to_keep)),
                as.numeric(names(wb4Style$rowHeights[[index]]))
              ),
              heights = ifelse(length(wb4Style$rowHeights[[index]]) == 0,
                "auto",
                as.numeric(wb4Style$rowHeights[[index]])
              )
            )
          }
        }


        ## output_EXCsheets:: Other styles ----

        if (isTRUE(output_EXCstyle %in% values_ExcelTableStyleNames)) {
          for (i in output_EXCsheets) {
            df_i <- readxl::read_xlsx(input_path,
              sheet = i,
              col_names = TRUE,
              trim_ws = TRUE,
              col_types = "text"
            )
            df_i <- as.data.frame(df_i)
            colnames(df_i) <- gsub("^[...]", "Column ", colnames(df_i))
            colnames(df_i) <- gsub("\\n", " ", colnames(df_i))

            openxlsx::addWorksheet(
              wb = MY_wb_copy,
              sheetName = i
            )

            openxlsx::writeDataTable(
              wb = MY_wb_copy,
              sheet = i,
              x = df_i,
              colNames = TRUE,
              tableStyle = output_EXCstyle
            )
            openxlsx::freezePane(
              wb = MY_wb_copy,
              sheet = i,
              firstRow = TRUE
            )
            setColWidths(
              wb = MY_wb_copy,
              sheet = i,
              cols = c(1:length(colnames(filtered_dataaa))),
              widths = "auto"
            )
            #+ freeze first row ?
            rm(df_i)
          }
        }

        if (output_EXCstyle == "none" | (output_EXCstyle == "keep_style" && !Style_exist)) {
          for (i in output_EXCsheets) {
            df_i <- readxl::read_xlsx(
              path = input_path,
              sheet = i,
              col_names = TRUE,
              trim_ws = TRUE,
              col_types = NULL)
            df_i <- as.data.frame(df_i)
            colnames(df_i) <- gsub("[...].*", "", colnames(df_i))
            openxlsx::addWorksheet(wb = MY_wb_copy,
                                   sheetName = i)
            openxlsx::writeData(wb = MY_wb_copy,
                                sheet = i,
                                x = df_i,
                                colNames = TRUE)
            rm(df_i)
          }
        }
      }


      ## :: Save WB ----

      path_filenaaame <- paste0(temp_folder_path, "/", filenaaame)

      openxlsx::saveWorkbook(
        wb = MY_wb_copy,
        file = path_filenaaame,
        overwrite = TRUE,
        returnValue = TRUE
      )

      rm(MY_wb_copy)
    }


    ## : Create CSV files ----

    if (truevalue_outputformat == "csv") {
      data_filtered <- MY_file[filter_rows, index_columns_to_keep]

      filenaaame <- paste0(output_names1, "(", filter_by_name, ") ", output_names2, ".csv")
      filenaaame <- fs::path_sanitize(filenaaame, replacement = "-")
      path_filenaaame <- paste0(temp_folder_path, "/", filenaaame)

      if (!is.null(ShinyApp_zipfilename)) {
        shiny::incProgress(
          amount = 8.5 / nb_files_created / 12,
          detail = filenaaame
        )
      }

      write.table(data_filtered, path_filenaaame,
        row.names = FALSE, na = "",
        fileEncoding = "UTF-16LE",
        sep = CSVsep[1],
        dec = CSVsep[2]
      )
    }
  }

  if (!is.null(ShinyApp_zipfilename)) {
    Sys.sleep(0.25)
    shiny::incProgress(
      amount = 1 / 12,
      message = "Ready to download !",
      detail = " "
    )
  }



  ## ZIP files -----------

  #print(paste0("temp_folder_path :", temp_folder_path))
  print("-- ZIPPING OUTPUT FILES...")

  if (!is.null(ShinyApp_zipfilename)) {
    print(paste0("Checking ShinyApp_zipfilename :", ShinyApp_zipfilename))

    utils::zip(
      zipfile = ShinyApp_zipfilename,
      files = dir(temp_folder_path, full.names = TRUE),
      extras = "-j"
    )

    shiny::incProgress(
      amount = 0,
      message = "Compressing files into ZIP...",
      detail = NULL
    )
    shiny::incProgress(
      amount = 0.5 / 12,
      message = "Done !"
    )
  } else {
    value_zipfilename <- paste0(truevalue_zipfileame, ".zip")
    #print(paste0("value_zipfilename :", value_zipfilename))

    return(utils::zip(
      zipfile = value_zipfilename,
      files = dir(temp_folder_path, full.names = TRUE),
      extras = "-j"
    ))
  }

  print("-- ZIP FILE DOWNLOADED !")
  print("You can find your file on your working directory : use getwd() ")



  if (TRUE) {
    unlink(temp_folder_path, recursive = TRUE)

    print((" "))
    print("Thank you for using the package ExcelFilterSplit :)")
    print((" "))
  }
}

