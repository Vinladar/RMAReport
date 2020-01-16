library(data.table)
library(dplyr)
rma_edit <- function(rma, rmaDetails, outputString = "Default.xlsx") {
	rma[, c(2, 3, 4, 6, 10, 12, 13, 15, 16)] = ""
	rma[,1] = convertToDate(rma[, 1], origin = "1900-01-01")
	rma[,2] <- rmaDetails$SO.ID
	columns <- c("Date", "orig_SO", "ShipDate", "Specifier", "Customer", "Job_Name", "Summary", "Product_Family", "Qty", "Qty_On_Order", "Prob_Part", "Replaced", "Repl_SO", "RMA", "Comments", "Code")
	colnames(rma) <- columns
	summ <- as.data.table(strsplit(rma$Summary, " // "))
	summ <- summ[2]
	rma$Summary <- t(summ)
	reason <- as.data.table(rmaDetails$Reason.Code)
	colnames(reason) <- "reason"
	codes <- read.csv("Reasoncode.csv", header = TRUE)
	reason1 <- left_join(reason, codes, by = "reason")
	rma$Code <- reason1[2]
	SO <- rma$orig_SO
	SO <- as.data.table(SO)
	SO[which(is.na(SO)),] <- ""
	rma$orig_SO <- SO
	rma <- as.data.table(rma)
	write.xlsx(rma, outputString, asTable = FALSE)
	return(rma)	
}