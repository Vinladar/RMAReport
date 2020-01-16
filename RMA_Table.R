# Import the required packages
library(dplyr)
library(data.table)
library(openxlsx)

# 
loadLightEngines <- function(fileName, output) {
	parseLightEngineStr <- function(engine) {
		return(substring(engine, 1, 7))
	}

	X <- read.xlsx(fileName)
	X <- X[,c("Item.ID", "Return.Qty")]
	light_engines <- X[grep("LEM-[0-9]{3}-[0-9]{2}", X$Item.ID),]
	engines <- as.data.table(apply(light_engines, 2, parseLightEngineStr))
	engines$Return.Qty <- as.numeric(engines$Return.Qty)
	engines <- aggregate(engines$Return.Qty, by = list(engines$Item.ID), sum)
	colnames(engines) <- c("Engine", "Quantity")
	engines <- engines[order(-engines$Quantity),]
	write.xlsx(engines, output, asTable=FALSE)
	return(engines)
}

reasonCodes <- function(fileName, output) {
	X <- read.xlsx(fileName)
	X2 <- as.data.table(X[, c("Reason.Code", "Return.Qty")])
	X2$Return.Qty <- as.numeric(X2$Return.Qty)
	X3 <- aggregate(X2$Return.Qty, by = list(X2$Reason.Code), sum)
	names(X3) <- c("Reason_Code", "Quantity")
	X4 <- X3[!is.na(X3$Quantity),]
	X4 <- X4[order(-X4$Quantity),]
	write.xlsx(X4, output, asTable = FALSE)
}

loadDrivers <- function(fileName, output) {
	X <- read.xlsx(fileName)
	X <- X[,c("Item.ID", "Return.Qty")]
	drivers <- X[grep("SP-[0-9]{3}-[0-9]{4}", X$Item.ID),]
	drivers[grep("EML$", drivers$Item.ID),1] <- substring(drivers[grep("EML$", drivers$Item.ID),1], 1, 13)
	substring(drivers[,1], 8, 11) <- "XXXX"
	drivers$Return.Qty <- as.numeric(drivers$Return.Qty)
	drivers <- as.data.table(aggregate(drivers$Return.Qty, by = list(drivers$Item.ID), sum))
	names(drivers) <- c("Driver", "Quantity")
	drivers <- drivers[order(-drivers$Quantity),]
	drivers <- filter(drivers, drivers$Quantity != 0)
	write.xlsx(drivers, output, asTable = FALSE)
}

generateReports <- function(fileName, lightEngines, drivers, codes) {
	loadLightEngines(fileName, lightEngines)
	loadDrivers(fileName, drivers)
	reasonCodes(fileName, codes)
}





















