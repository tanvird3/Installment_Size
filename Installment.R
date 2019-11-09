library(openxlsx)
library(plyr)
library(data.table)
library(dplyr)
library(tidyr)

x <- dirname(rstudioapi::getSourceEditorContext()$path)

setwd(x)

date <- as.Date("2019-10-01", format = "%Y-%m-%d")

### MDS ###

dx <- read.xlsx("mds.xlsx", detectDates = T)

dx <- mutate(dx, CB = Current.Balance / 10000000)

d <- filter(dx,!is.na(Rate) & Rate %in% c(7.80, 8.72, 7.27, 8.08))

d$ACCOUNT_TYPE <-
  mapvalues(
    d$Rate,
    from = c(7.80, 8.72, 7.27, 8.08),
    to = c("OLD CIRCULAR", "OLD CIRCULAR", "NEW CIRCULAR", "NEW CIRCULAR")
  )

d$Months_Elapsed <- floor(as.numeric(date - d$Open.Date) / 30)

d$Installment_Size <- d$Current.Balance / d$Months_Elapsed

d$Installment_Size[!is.finite(d$Installment_Size)] <-
  d$Current.Balance[!is.finite(d$Installment_Size)]

lev <-
  c(0, 500, 1000, 2000, 5000, 10000, 15000, 20000, 25000, 30000, Inf)

lab <-
  c(500, 500, 1000, 2000, 5000, 10000, 15000, 20000, 25000, 30000)

d$Installment_Size <-
  as.numeric(as.character(cut(d$Installment_Size, lev, lab, right = F)))


mds_summary <-
  data.frame(
    d %>% group_by(ACCOUNT_TYPE, Product.Name, Rate) %>% summarise(
      AC_NUMBER = length(Current.Balance),
      Total_Deposit = sum(Installment_Size)
    )
  )

if (as.vector(Sys.info()['sysname']) == "Windows") {
  Sys.setenv("R_ZIPCMD" = "C:/Rtools/bin/zip.exe")
}

write.xlsx(mds_summary,
           "MDS_INST.xlsx",
           asTable = F,
           row.names = F)

#### scheme #####

dy <- read.xlsx("scheme.xlsx", detectDates = T)

d <-
  filter(
    dy,!is.na(Rate) &
      Rate %in% c(7.45, 7.80, 7.00, 7.27) &
      Open.Date >= as.Date("2018-08-12", format = "%Y-%m-%d")
  )

d$ACCOUNT_TYPE <-
  mapvalues(
    d$Rate,
    from = c(7.45, 7.80, 7.00, 7.27),
    to = c("OLD CIRCULAR", "OLD CIRCULAR", "NEW CIRCULAR", "NEW CIRCULAR")
  )

d$Months_Elapsed <- floor(as.numeric(date - d$Open.Date) / 30)

d$Installment_Size <- d$Current.Balance / d$Months_Elapsed

d$Installment_Size[!is.finite(d$Installment_Size)] <-
  d$Current.Balance[!is.finite(d$Installment_Size)]

lev <-
  c(0, 500, 1000, 2000, 5000, 10000, 15000, 20000, 25000, 30000, Inf)

lab <-
  c(500, 500, 1000, 2000, 5000, 10000, 15000, 20000, 25000, 30000)

d$Installment_Size <-
  as.numeric(as.character(cut(d$Installment_Size, lev, lab, right = F)))


scheme_summary <-
  data.frame(
    d %>% group_by(ACCOUNT_TYPE, Product.Name, Rate) %>% summarise(
      AC_NUMBER = length(Current.Balance),
      Total_Deposit = sum(Installment_Size)
    )
  )

write.xlsx(scheme_summary,
           "SCHEME_INST.xlsx",
           asTable = F,
           row.names = F)

## millionaire ###

mn <- read.xlsx("miln.xlsx", detectDates = T)

mn <- filter(mn, Rate %in% c(8.30, 8.00))

mn$AC <- paste0(mn$Branch.ID, mn$Account.Number)

mn <- select(mn, 15, 5, 6, 7)


mo <- read.xlsx("milo.xlsx", detectDates = T)

mo <- filter(mo, Rate %in% c(8.30, 8.00))

mo$AC <- paste0(mo$Branch.ID, mo$Account.Number)

mo <- select(mo, 15, 7)


ml <- left_join(mn, mo, by = "AC")

ml$Current.Balance.x[is.na(ml$Current.Balance.x)] <- 0
ml$Current.Balance.y[is.na(ml$Current.Balance.y)] <- 0

ml$Installment_Size <- ml$Current.Balance.x - ml$Current.Balance.y


size <- read.xlsx("milinstsize.xlsx")

size <- size %>% distinct(size, .keep_all = T)

size <- size %>% arrange(size)

lev <- c(0, size$size, Inf)

lab <- c(size$size[1], size$size)

ml$Installment_Size <-
  as.numeric(as.character(cut(ml$Installment_Size, lev, lab, right = F)))

ml$ACCOUNT_TYPE <-
  mapvalues(ml$Rate,
            from = c(8.30, 8.00),
            to = c("OLD CIRCULAR", "NEW CIRCULAR"))

mill_summary <-
  data.frame(
    ml %>% group_by(ACCOUNT_TYPE, Product.Name, Rate) %>% summarise(
      AC_NUMBER = length(Current.Balance.x),
      Total_Deposit = sum(Installment_Size)
    )
  )

write.xlsx(mill_summary,
           "mill_INST.xlsx",
           asTable = F,
           row.names = F)
