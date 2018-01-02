### 01_Main_Execution_Script.R
# --Parsing SAP Ticket Files - Weekly APGI Reporting
# --Count of new tickets created per month, by month and average ticket outstanding time, by month

# Load libraries
  library(readr)       # Load CSV into data.frame
  library(lubridate)   # Manipulation of dates
  library(reshape2)    # Melting data.frame for GGPlot
  library(ggplot2)     # Plotting

# Import Excel File
  SAP_REPORT <- read_csv("~/Desktop/Stuff/R/Weekly_Reports/&__Data/SAP-REPORT.csv")
  names(SAP_REPORT) <- c("Type","WorkCtr","Notification","Created","Completed","Priority","PrText","PrType","WorkOrder","SortField",
                         "ShortDesc","CreatedBy","LongDesc","ReportedBy","Status")

#Clean up files - dates to dates, numbers to numbers
  SAP_REPORT$Notification <- as.numeric(SAP_REPORT$Notification)
  SAP_REPORT$Created <- as.Date(SAP_REPORT$Created,"%m/%d/%Y")
  SAP_REPORT$Completed <- as.Date(SAP_REPORT$Completed,"%m/%d/%Y")
  SAP_REPORT$Priority <- as.numeric(SAP_REPORT$Priority)
  SAP_REPORT$WorkOrder <- as.numeric(SAP_REPORT$WorkOrder)
  
# Create GK and VM specific DFs, and strip erroneous data
  SAP_GK <- subset(data.frame(SAP_REPORT),WorkCtr == 'GKEST')
  SAP_VM <- subset(data.frame(SAP_REPORT),WorkCtr == 'VSEE1')
  
#Create parsing tables
  Parsing_ColNames <- c("Created","Month","Year","Completed","Outstanding","Priority","Status")

  Parsing_GK <- data.frame(matrix(ncol=7,nrow=nrow(SAP_GK)))
  Parsing_VM <- data.frame(matrix(ncol=7,nrow=nrow(SAP_VM)))

  colnames(Parsing_GK) <- Parsing_ColNames
  colnames(Parsing_VM) <- Parsing_ColNames

  rm(Parsing_ColNames)

# Migrate data to parsing tables - Columns 1, 2, 3, 4, 6
  Parsing_GK$Created <- SAP_GK$Created
  Parsing_GK$Month <- month(SAP_GK$Created)
  Parsing_GK$Year <- year(SAP_GK$Created)
  Parsing_GK$Completed <- SAP_GK$Completed
  Parsing_GK$Priority <- SAP_GK$Priority


  Parsing_VM$Created <- SAP_VM$Created
  Parsing_VM$Month <- month(SAP_VM$Created)
  Parsing_VM$Year <- year(SAP_VM$Created)
  Parsing_VM$Completed <- SAP_VM$Completed
  Parsing_VM$Priority <- SAP_VM$Priority


## Parsing Data
# If no completion date, add todays date
  Parsing_GK$Completed[is.na(Parsing_GK$Completed)] <- Sys.Date()
  Parsing_VM$Completed[is.na(Parsing_VM$Completed)] <- Sys.Date()
  
# Add outstanding days - Column 5
  Parsing_GK$Outstanding <- as.numeric(Parsing_GK$Completed - Parsing_GK$Created)
  Parsing_VM$Outstanding <- as.numeric(Parsing_VM$Completed - Parsing_VM$Created)

# Add status - Column 7
  Counter <- c(1)

  while(Counter != nrow(Parsing_GK) + 1)
  {
    if(grepl("NOCO",SAP_GK$Status[Counter]))
    {
    Parsing_GK$Status[Counter] <- "Closed"
    }
    
    if(grepl("OSNO",SAP_GK$Status[Counter]))
    {
      Parsing_GK$Status[Counter] <- "Outstanding"
    }

    if(grepl("NOPR", SAP_GK$Status[Counter]))
    {
      Parsing_GK$Status[Counter] <- "Work Order"
    }
  
    Counter <- Counter + 1
  }

  Counter <- c(1)

  while(Counter != nrow(Parsing_VM) + 1)
  {
    if(grepl("NOCO",SAP_VM$Status[Counter]))
    {
      Parsing_VM$Status[Counter] <- "Closed"
    }
    
    if(grepl("OSNO",SAP_VM$Status[Counter]))
    {
      Parsing_VM$Status[Counter] <- "Outstanding"
    }
  
    if(grepl("NOPR", SAP_VM$Status[Counter]))
    {
      Parsing_VM$Status[Counter] <- "Work Order"
    }
  
    Counter <- Counter + 1
  }

## Extract and calculate current record set - Last six months data
# Get record start date-first of the month
  StartDate <- Sys.Date() %m-% months(5)
  day(StartDate) <- 01

# Delete all records older than start date - not required (month and year checked later)
  Parsing_GK <- Parsing_GK[!(Parsing_GK$Created < StartDate),]
  Parsing_VM <- Parsing_VM[!(Parsing_VM$Created < StartDate),]

# Create DFs with names
  LastSixMonths_GK <-data.frame(matrix(ncol=12,nrow=6))
  LastSixMonths_VM <-data.frame(matrix(ncol=12,nrow=6))

  ColN <- c("Month","Year","Priority 1 - Total Notifications","Priority 1 - Average Outstanding",
            "Priority 2 - Total Notifications","Priority 2 - Average Outstanding",
            "Priority 3 - Total Notifications","Priority 3 - Average Outstanding",
            "Priority 4 - Total Notifications","Priority 4 - Average Outstanding",
            "Priority 5 - Total Notifications","Priority 5 - Average Outstanding")
  
  colnames(LastSixMonths_GK) <- ColN
  colnames(LastSixMonths_VM) <- ColN
  
# Add last six month - Month and year
  LastSixMonths_GK$Month[1] <- month(StartDate); LastSixMonths_GK$Year <- year(StartDate)
  LastSixMonths_GK$Month[2] <- month(StartDate %m+% months(1)); LastSixMonths_GK$Year <-year(StartDate %m+% months(1))
  LastSixMonths_GK$Month[3] <- month(StartDate %m+% months(2)); LastSixMonths_GK$Year <-year(StartDate %m+% months(2))
  LastSixMonths_GK$Month[4] <- month(StartDate %m+% months(3)); LastSixMonths_GK$Year <-year(StartDate %m+% months(3))
  LastSixMonths_GK$Month[5] <- month(StartDate %m+% months(4)); LastSixMonths_GK$Year <-year(StartDate %m+% months(4))
  LastSixMonths_GK$Month[6] <- month(StartDate %m+% months(5)); LastSixMonths_GK$Year <-year(StartDate %m+% months(5))

  LastSixMonths_VM$Month <- LastSixMonths_GK$Month; LastSixMonths_VM$Year <- LastSixMonths_GK$Year;

# Load monthly totals
  Counter <- c(1)
  
  while(Counter < 7)
  {
    LastSixMonths_GK[Counter,3] <- length(Parsing_GK[(Parsing_GK$Priority == 1 & month(Parsing_GK$Created) == LastSixMonths_GK[Counter,1] & year(Parsing_GK$Created) == LastSixMonths_GK[Counter,2]),4])
    LastSixMonths_GK[Counter,5] <- length(Parsing_GK[(Parsing_GK$Priority == 2 & month(Parsing_GK$Created) == LastSixMonths_GK[Counter,1] & year(Parsing_GK$Created) == LastSixMonths_GK[Counter,2]),4])
    LastSixMonths_GK[Counter,7] <- length(Parsing_GK[(Parsing_GK$Priority == 3 & month(Parsing_GK$Created) == LastSixMonths_GK[Counter,1] & year(Parsing_GK$Created) == LastSixMonths_GK[Counter,2]),4])
    LastSixMonths_GK[Counter,9] <- length(Parsing_GK[(Parsing_GK$Priority == 4 & month(Parsing_GK$Created) == LastSixMonths_GK[Counter,1] & year(Parsing_GK$Created) == LastSixMonths_GK[Counter,2]),4])
    LastSixMonths_GK[Counter,11] <- length(Parsing_GK[(Parsing_GK$Priority == 5 & month(Parsing_GK$Created) == LastSixMonths_GK[Counter,1] & year(Parsing_GK$Created) == LastSixMonths_GK[Counter,2]),4])
    
    Counter <- Counter + 1
  }
  
  Counter <- c(1)
  
  while(Counter < 7)
  {
    LastSixMonths_VM[Counter,3] <- length(Parsing_VM[(Parsing_VM$Priority == 1 & month(Parsing_VM$Created) == LastSixMonths_VM[Counter,1] & year(Parsing_VM$Created) == LastSixMonths_VM[Counter,2]),4])
    LastSixMonths_VM[Counter,5] <- length(Parsing_VM[(Parsing_VM$Priority == 2 & month(Parsing_VM$Created) == LastSixMonths_VM[Counter,1] & year(Parsing_VM$Created) == LastSixMonths_VM[Counter,2]),4])
    LastSixMonths_VM[Counter,7] <- length(Parsing_VM[(Parsing_VM$Priority == 3 & month(Parsing_VM$Created) == LastSixMonths_VM[Counter,1] & year(Parsing_VM$Created) == LastSixMonths_VM[Counter,2]),4])
    LastSixMonths_VM[Counter,9] <- length(Parsing_VM[(Parsing_VM$Priority == 4 & month(Parsing_VM$Created) == LastSixMonths_VM[Counter,1] & year(Parsing_VM$Created) == LastSixMonths_VM[Counter,2]),4])
    LastSixMonths_VM[Counter,11] <- length(Parsing_VM[(Parsing_VM$Priority == 5 & month(Parsing_VM$Created) == LastSixMonths_VM[Counter,1] & year(Parsing_VM$Created) == LastSixMonths_VM[Counter,2]),4])
    
    Counter <- Counter + 1
  }
  
# Load Outstanding
  A <- 1
  
  while(A < 7)
  {
    if(is.na(sum(Parsing_GK[which(Parsing_GK$Priority == 1 & month(Parsing_GK$Created) == LastSixMonths_GK[A,1] & year(Parsing_GK$Created) == LastSixMonths_GK[A,2]),5])))
    {      
      LastSixMonths_GK[A,4] <- 0
    }
    else
    {
      LastSixMonths_GK[A,4] <- round(sum(Parsing_GK[which(Parsing_GK$Priority == 1 & month(Parsing_GK$Created) == LastSixMonths_GK[A,1] & year(Parsing_GK$Created) == LastSixMonths_GK[A,2]),5]) / LastSixMonths_GK[A,3],0)
    }
    
    if(is.na(sum(Parsing_GK[which(Parsing_GK$Priority == 2 & month(Parsing_GK$Created) == LastSixMonths_GK[A,1] &
                                  year(Parsing_GK$Created) == LastSixMonths_GK[A,2]),3])))
    {
      LastSixMonths_GK[A,6] <- 0
    }
    else
    {
      LastSixMonths_GK[A,6] <- round(sum(Parsing_GK[which(Parsing_GK$Priority == 2 & month(Parsing_GK$Created) == LastSixMonths_GK[A,1] &
                                  year(Parsing_GK$Created) == LastSixMonths_GK[A,2]),5]) / LastSixMonths_GK[A,5],0)
    }
    
    if(is.na(sum(Parsing_GK[which(Parsing_GK$Priority == 3 & month(Parsing_GK$Created) == LastSixMonths_GK[A,1] &
                                  year(Parsing_GK$Created) == LastSixMonths_GK[A,2]),3])))
    {
      LastSixMonths_GK[A,8] <- 0
    }
    else
    {
      LastSixMonths_GK[A,8] <- round(sum(Parsing_GK[which(Parsing_GK$Priority == 3 & month(Parsing_GK$Created) == LastSixMonths_GK[A,1] & year(Parsing_GK$Created) == LastSixMonths_GK[A,2]),5]) / LastSixMonths_GK[A,7],0)
    } 
    
    if(is.na(sum(Parsing_GK[which(Parsing_GK$Priority == 4 & month(Parsing_GK$Created) == LastSixMonths_GK[A,1] & year(Parsing_GK$Created) == LastSixMonths_GK[A,2]),3])))
    {
      LastSixMonths_GK[A,10] <- 0
    }
    else
    {
      LastSixMonths_GK[A,10] <- round(sum(Parsing_GK[which(Parsing_GK$Priority == 4 & month(Parsing_GK$Created) == LastSixMonths_GK[A,1] & year(Parsing_GK$Created) == LastSixMonths_GK[A,2]),5]) / LastSixMonths_GK[A,9],0)
    }
    
    if(is.na(sum(Parsing_GK[which(Parsing_GK$Priority == 5 & month(Parsing_GK$Created) == LastSixMonths_GK[A,1] & year(Parsing_GK$Created) == LastSixMonths_GK[A,2]),3])))
    {
      LastSixMonths_GK[A,12] <- 0
    }
    else
    {
      LastSixMonths_GK[A,12] <- round(sum(Parsing_GK[which(Parsing_GK$Priority == 5 & month(Parsing_GK$Created) == LastSixMonths_GK[A,1] & year(Parsing_GK$Created) == LastSixMonths_GK[A,2]),5]) / LastSixMonths_GK[A,11],0)
    }
    
    A <- A + 1
  }
  
  A <- 1
  
  while(A < 7)
  {
    if(is.na(sum(Parsing_VM[which(Parsing_VM$Priority == 1 & month(Parsing_VM$Created) == LastSixMonths_VM[A,1] & year(Parsing_VM$Created) == LastSixMonths_VM[A,2]),3])))
    {
      LastSixMonths_VM[A,4] <- 0
    }
    else
    {
      LastSixMonths_VM[A,4] <- round(sum(Parsing_VM[which(Parsing_VM$Priority == 1 & month(Parsing_VM$Created) == LastSixMonths_VM[A,1] & year(Parsing_VM$Created) == LastSixMonths_VM[A,2]),5]) / LastSixMonths_VM[A,3],0)
    }
    
    if(is.na(sum(Parsing_VM[which(Parsing_VM$Priority == 2 & month(Parsing_VM$Created) == LastSixMonths_VM[A,1] & year(Parsing_VM$Created) == LastSixMonths_VM[A,2]),3])))
    {
      LastSixMonths_VM[A,6] <- 0
    }
    else
    {
      LastSixMonths_VM[A,6] <- round(sum(Parsing_VM[which(Parsing_VM$Priority == 2 & month(Parsing_VM$Created) == LastSixMonths_VM[A,1] & year(Parsing_VM$Created) == LastSixMonths_VM[A,2]),5]) / LastSixMonths_VM[A,5],0)
    }
    
    if(is.na(sum(Parsing_VM[which(Parsing_VM$Priority == 3 & month(Parsing_VM$Created) == LastSixMonths_VM[A,1] & year(Parsing_VM$Created) == LastSixMonths_VM[A,2]),3])))
    {
      LastSixMonths_VM[A,8] <- 0
    }
    else
    {
      LastSixMonths_VM[A,8] <- round(sum(Parsing_VM[which(Parsing_VM$Priority == 3 & month(Parsing_VM$Created) == LastSixMonths_VM[A,1] & year(Parsing_VM$Created) == LastSixMonths_VM[A,2]),5]) / LastSixMonths_VM[A,7],0)
    } 
    
    if(is.na(sum(Parsing_VM[which(Parsing_VM$Priority == 4 & month(Parsing_VM$Created) == LastSixMonths_VM[A,1] & year(Parsing_VM$Created) == LastSixMonths_VM[A,2]),3])))
    {
      LastSixMonths_VM[A,10] <- 0
    }
    else
    {
      LastSixMonths_VM[A,10] <- round(sum(Parsing_VM[which(Parsing_VM$Priority == 4 & month(Parsing_VM$Created) == LastSixMonths_VM[A,1] & year(Parsing_VM$Created) == LastSixMonths_VM[A,2]),5]) / LastSixMonths_VM[A,9],0)
    }
    
    if(is.na(sum(Parsing_VM[which(Parsing_VM$Priority == 5 & month(Parsing_VM$Created) == LastSixMonths_VM[A,1] & year(Parsing_VM$Created) == LastSixMonths_VM[A,2]),3])))
    {
      LastSixMonths_VM[A,12] <- 0
    }
    else
    {
      LastSixMonths_VM[A,12] <- round(sum(Parsing_VM[which(Parsing_VM$Priority == 5 & month(Parsing_VM$Created) == LastSixMonths_VM[A,1] & year(Parsing_VM$Created) == LastSixMonths_VM[A,2]),5]) / LastSixMonths_VM[A,11],0)
    }
    
    A <- A + 1
  }
  
# Change any NaN (divide by 0) to 0
  LastSixMonths_GK <- replace(LastSixMonths_GK,is.na(LastSixMonths_GK),0) ; LastSixMonths_VM <- replace(LastSixMonths_VM,is.na(LastSixMonths_VM),0)
  
# Create Plotting Tables
  LastSix_GK <- data.frame() ; LastSix_VM <- data.frame()
  LastSix_GK <- LastSixMonths_GK ; LastSix_VM <- LastSixMonths_VM
  LastSix_GK$Month <- paste(LastSix_GK$Month," - ",LastSix_GK$Year) ; LastSix_VM$Month <- paste(LastSix_VM$Month," - ",LastSix_VM$Year)
  LastSix_GK <- LastSix_GK[,-2] ; LastSix_VM <- LastSix_VM[,-2]
  
  ### Plotting All
  # Y Scale - Find Highest # in each of GK and VM Frames
  GK_Scale <- max(LastSix_GK[2:11], na.rm=TRUE) ; VM_Scale <- max(LastSix_VM[2:11], na.rm=TRUE)
  
  # Round up to nearest 10
  GK_Scale <- ceiling(GK_Scale/10)*10 ; VM_Scale <- ceiling(VM_Scale/10)*10
  
  # Melt data to create long vs wide
  df_G <- melt(LastSix_GK,id.vars = "Month") ; df_V <- melt(LastSix_VM,id.vars = "Month")
  
  # Plot Parameters
  MonOrderG <- c(LastSix_GK[1,1],LastSix_GK[2,1],LastSix_GK[3,1],LastSix_GK[4,1],LastSix_GK[5,1],LastSix_GK[6,1])
  MonOrderV <- c(LastSix_VM[1,1],LastSix_VM[2,1],LastSix_VM[3,1],LastSix_VM[4,1],LastSix_VM[5,1],LastSix_VM[6,1])
  GK_Title <- c(paste("Gahcho Kue Mine - SAP Notifications - Last 6 Months - report date:",format(Sys.Date(),"%b-%d-%Y")))
  VM_Title <- c(paste("Victor Mine - SAP Notifications - Last 6 Months - report date:",format(Sys.Date(),"%b-%d-%Y")))
   
  #Colour Pallette
  cbPalette <- c("#000000","#E69F00","#666666","#56B4E9","#999999","#009E73","#cccccc","#F0E442","#ffffff","#0072B2","#D55E00","#CC79A7")
  
  #Save to PDF Wrapper
  pdfname = c(paste("~/Desktop/Stuff/R/Weekly_Reports/&__Data/",format(Sys.Date(),"%m-%d-%Y"),"- PS Weekly Reporting.pdf"))
  pdf(pdfname, height = 5.5, width = 8.5)
  
  #Plot Gahcho to Wrapper
  ggplot(df_G, aes(x=Month, y=value, fill=variable)) + scale_x_discrete(limits = MonOrderG) + 
    geom_bar(stat='identity', position ='dodge', colour = 'black') + xlab("Month/Year") + ylab("New Notifications/Average Outstanding") + 
    ggtitle(GK_Title) + scale_fill_manual(values=cbPalette) + labs(fill = "Legend") +
    scale_y_continuous(limits = c(0,GK_Scale), breaks = seq(0,GK_Scale,10), minor_breaks = seq(10,110,20))

  #Plot Victor to Wrapper
    ggplot(df_V, aes(x=Month, y=value, fill=variable)) + scale_x_discrete(limits = MonOrderV) +  
    geom_bar(stat='identity', position ='dodge', colour = 'black') + xlab("Month/Year") + ylab("New Notifications/Average Outstanding") +
    ggtitle(VM_Title) + scale_fill_manual(values=cbPalette) + labs(fill = "Legend") +
    scale_y_continuous(limits = c(0,VM_Scale), breaks = seq(0,VM_Scale,10), minor_breaks = seq(10,110,20))

  #Close PDF Wrapper
  dev.off()

  #Plot Gahcho to Plots
  ggplot(df_G, aes(x=Month, y=value, fill=variable)) + scale_x_discrete(limits = MonOrderG) + 
      geom_bar(stat='identity', position ='dodge', colour = 'black') + xlab("Month/Year") + ylab("New Notifications/Average Outstanding") + 
      ggtitle(GK_Title) + scale_fill_manual(values=cbPalette) + labs(fill = "Legend") +
      scale_y_continuous(limits = c(0,GK_Scale), breaks = seq(0,GK_Scale,10), minor_breaks = seq(10,110,20))
  
  #Plot Victor to Plots
  ggplot(df_V, aes(x=Month, y=value, fill=variable)) + scale_x_discrete(limits = MonOrderV) +  
      geom_bar(stat='identity', position ='dodge', colour = 'black') + xlab("Month/Year") + ylab("New Notifications/Average Outstanding") +
      ggtitle(VM_Title) + scale_fill_manual(values=cbPalette) + labs(fill = "Legend") +
      scale_y_continuous(limits = c(0,VM_Scale), breaks = seq(0,VM_Scale,10), minor_breaks = seq(10,110,20))

  