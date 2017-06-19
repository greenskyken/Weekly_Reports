### 01_Main_Execution_Script.R
# --Parsing SAP Ticket Files - Weekly APGI Reporting
# --Count of new tickets created per month, by month
# --Average ticket outstanding time, by month

# Set Working Directory
	setwd("C:/Users/ken.bartlett/Desktop/My Stuff/R/Weekly_Reports")

# Import Excel File
	library(readr)
  SAP_CSV <- read_csv("C:/Users/ken.bartlett/Desktop/My Stuff/R/Weekly_Reports/&__Data/SAP-REPORT.csv")
	A_SAP<- data.frame(SAP_CSV)

# Clean all but GK & VM TSS Records
	DF_SAP_GK <- subset(A_SAP,Main.WorkCtr == "GKEST")
	DF_SAP_VM <- subset(A_SAP,Main.WorkCtr == "VSEE1")

### Create Parsing - Gahcho Kue
# Create Column Names
	ColNames <- c("GKEST","Notification","Order","Created","Completed","Outstanding","Month","Year","Priority","Status","Closed","Open")

# Get Row Counts
	GK_RowCount <- nrow(DF_SAP_GK)

# Create Parsing Tables
	GK_Parsing <- data.frame()

# Load Static Items
	Counter_GK <- 1
	
	while(Counter_GK != GK_RowCount+1)
	{
	  GK_Parsing[Counter_GK,1] <- DF_SAP_GK[Counter_GK,2]
	  GK_Parsing[Counter_GK,2] <- DF_SAP_GK[Counter_GK,3]
	  GK_Parsing[Counter_GK,3] <- DF_SAP_GK[Counter_GK,9]
	  GK_Parsing[Counter_GK,4] <- DF_SAP_GK[Counter_GK,4]
	  GK_Parsing[Counter_GK,5] <- DF_SAP_GK[Counter_GK,5]
	  GK_Parsing[Counter_GK,6] <- c(0) 
	  GK_Parsing[Counter_GK,7] <- c(0)   
	  GK_Parsing[Counter_GK,8] <- c(0)   
	  GK_Parsing[Counter_GK,9] <- c(0)   
	  GK_Parsing[Counter_GK,10] <- c(0)   
	  GK_Parsing[Counter_GK,11] <- c(0)   
	  GK_Parsing[Counter_GK,12] <- c(0)   
	  
	  Counter_GK <- Counter_GK+1 
	}

	colnames(GK_Parsing) <- c(ColNames)

# Load Outstanding
	Counter_GK <- 1
	Temp_1 <- 0

	while(Counter_GK != GK_RowCount+1)
	{
	  Temp_1 <- as.Date(GK_Parsing[Counter_GK,5], origin="1900-01-01",format="%m/%d/%y") -
		as.Date(GK_Parsing[Counter_GK,4], origin="1900-01-01",format="%m/%d/%y")
	  
	  if(is.na(as.Date(GK_Parsing[Counter_GK,5], origin="1900-01-01",format="%m/%d/%y") -
			   as.Date(GK_Parsing[Counter_GK,4], origin="1900-01-01",format="%m/%d/%y")))
	  {
		X <- Sys.Date()
		Y <- as.Date(GK_Parsing[Counter_GK,4],format="%m/%d/%Y")
		
		GK_Parsing[Counter_GK,6] <- c(X-Y)
	  }
	  else if(Temp_1 < 0)
	  {
		GK_Parsing[Counter_GK,6] <- Temp_1 + 366
	  }
	  else if(Temp_1 > 0)
	  {
		GK_Parsing[Counter_GK,6] <- Temp_1
	  }
	  
	  Counter_GK <- Counter_GK+1 
	}

# Load Month and Year
	Counter_GK <- 1
	Month_GK <- 0
	Year_GK <- 0

	while(Counter_GK != GK_RowCount+1)
	{
	  Date1 <- GK_Parsing[Counter_GK,4]
	  Len <- nchar(GK_Parsing[Counter_GK,4])
	  
	  Year1 <- substring(Date1,Len-3,Len)
	  
	  if(substring(Date1,2,2) == "/")
	  {
		Month1 <- as.numeric(substring(Date1,1,1)) 
	  }
	  else
	  {
		Month1 <- as.numeric(substring(Date1,1,2))
	  }
	  
	  GK_Parsing[Counter_GK,7] <- Month1
	  GK_Parsing[Counter_GK,8] <- Year1
	  
	  Counter_GK <- Counter_GK+1 
	}

# Load Priority
	Counter_GK <- 1

	while(Counter_GK != GK_RowCount+1)
	{
	  GK_Parsing[Counter_GK,9] <- DF_SAP_GK[Counter_GK,6]
	  
	  Counter_GK <- Counter_GK+1 
	}

# Load Status
	Counter_GK <- 1

	while(Counter_GK != GK_RowCount+1)
	{  
	  SAPStatus <- DF_SAP_GK[Counter_GK,15]
	  Status1 <- 0
	  
	  if(grepl("NOCO",SAPStatus))
	  {
		Status1 <- "CLOSED" 
	  }
	  
	  if(grepl("OSNO",SAPStatus))
	  {
		Status1 <- "OUTSTANDING"
	  }
	  
	  if(grepl("NOPR",SAPStatus))
	  {
		Status1 <- "WORK ORDER"
	  }
	  
	  GK_Parsing[Counter_GK,10] <- Status1
	  
	  Counter_GK <- Counter_GK+1
	}

# Load Opened/Closed
	Counter_GK <- 1

	while(Counter_GK != GK_RowCount+1)
	{  
	  if(GK_Parsing[Counter_GK,10] == "CLOSED")
	  {
		GK_Parsing[Counter_GK,11] <- 1
		GK_Parsing[Counter_GK,12] <- 0
	  }
	  else
	  {
		GK_Parsing[Counter_GK,11] <- 0
		GK_Parsing[Counter_GK,12] <- 1
	  }
	  Counter_GK <- Counter_GK+1
	}

### Create Parsing - Victor Mine

# Create Column Names
	ColNames <- c("VSEE1","Notification","Order","Created","Completed","Outstanding","Month","Year","Priority","Status","Closed","Open")

# Get Row Counts
	VM_RowCount <- nrow(DF_SAP_VM)

# Create Parsing Tables
	VM_Parsing <- data.frame()

# Load Static Items
	Counter_VM <- 1

	while(Counter_VM != VM_RowCount+1)
	{
	  VM_Parsing[Counter_VM,1] <- DF_SAP_VM[Counter_VM,2]
	  VM_Parsing[Counter_VM,2] <- DF_SAP_VM[Counter_VM,3]
	  VM_Parsing[Counter_VM,3] <- DF_SAP_VM[Counter_VM,9]
	  VM_Parsing[Counter_VM,4] <- DF_SAP_VM[Counter_VM,4]
	  VM_Parsing[Counter_VM,5] <- DF_SAP_VM[Counter_VM,5]
	  VM_Parsing[Counter_VM,6] <- c(0) 
	  VM_Parsing[Counter_VM,7] <- c(0)   
	  VM_Parsing[Counter_VM,8] <- c(0)   
	  VM_Parsing[Counter_VM,9] <- c(0)   
	  VM_Parsing[Counter_VM,10] <- c(0)   
	  VM_Parsing[Counter_VM,11] <- c(0)   
	  VM_Parsing[Counter_VM,12] <- c(0)   
	  
	  Counter_VM <- Counter_VM+1 
	}

	colnames(VM_Parsing) <- c(ColNames)

# Load Outstanding
	Counter_VM <- 1
	Temp_1 <- 0

	while(Counter_VM != VM_RowCount+1)
	{
	  Temp_1 <- as.Date(VM_Parsing[Counter_VM,5], origin="1900-01-01",format="%m/%d/%y") -
		as.Date(VM_Parsing[Counter_VM,4], origin="1900-01-01",format="%m/%d/%y")
	  
	  if(is.na(as.Date(VM_Parsing[Counter_VM,5], origin="1900-01-01",format="%m/%d/%y") -
			   as.Date(VM_Parsing[Counter_VM,4], origin="1900-01-01",format="%m/%d/%y")))
	  {
		X <- Sys.Date()
		Y <- as.Date(VM_Parsing[Counter_VM,4],format="%m/%d/%Y")
		
		VM_Parsing[Counter_VM,6] <- c(X-Y)
	  }
	  else if(Temp_1 < 0)
	  {
		VM_Parsing[Counter_VM,6] <- Temp_1 + 366
	  }
	  else if(Temp_1 > 0)
	  {
		VM_Parsing[Counter_VM,6] <- Temp_1
	  }
	  
	  Counter_VM <- Counter_VM+1 
	}

# Load Month and Year
	Counter_VM <- 1
	Month_VM <- 0
	Year_VM <- 0

	while(Counter_VM != VM_RowCount+1)
	{
	  Date1 <- VM_Parsing[Counter_VM,4]
	  Len <- nchar(VM_Parsing[Counter_VM,4])
	  
	  Year1 <- substring(Date1,Len-3,Len)
	  
	  if(substring(Date1,2,2) == "/")
	  {
		Month1 <- as.numeric(substring(Date1,1,1)) 
	  }
	  else
	  {
		Month1 <- as.numeric(substring(Date1,1,2))
	  }
	  
	  VM_Parsing[Counter_VM,7] <- Month1
	  VM_Parsing[Counter_VM,8] <- Year1
	  
	  Counter_VM <- Counter_VM+1 
	}

# Load Priority
	Counter_VM <- 1

	while(Counter_VM != VM_RowCount+1)
	{
	  VM_Parsing[Counter_VM,9] <- DF_SAP_VM[Counter_VM,6]
	  
	  Counter_VM <- Counter_VM+1 
	}

# Load Status
	Counter_VM <- 1

	while(Counter_VM != VM_RowCount+1)
	{  
	  SAPStatus <- DF_SAP_VM[Counter_VM,15]
	  Status1 <- 0
	  
	  if(grepl("NOCO",SAPStatus))
	  {
		Status1 <- "CLOSED" 
	  }
	  
	  if(grepl("OSNO",SAPStatus))
	  {
		Status1 <- "OUTSTANDING"
	  }
	  
	  if(grepl("NOPR",SAPStatus))
	  {
		Status1 <- "WORK ORDER"
	  }
	  
	  VM_Parsing[Counter_VM,10] <- Status1
	  
	  Counter_VM <- Counter_VM+1
	}

# Load Opened/Closed
	Counter_VM <- 1

	while(Counter_VM != VM_RowCount+1)
	{  
	  if(VM_Parsing[Counter_VM,10] == "CLOSED")
	  {
		VM_Parsing[Counter_VM,11] <- 1
		VM_Parsing[Counter_VM,12] <- 0
	  }
	  else
	  {
		VM_Parsing[Counter_VM,11] <- 0
		VM_Parsing[Counter_VM,12] <- 1
	  }
	  Counter_VM <- Counter_VM+1
	}

### Create Plot Data - Gahcho Kue
	CurrDate <- Sys.Date()
	CurrMonthDate <- format(as.Date(CurrDate),"%m")
	CurrYearDate <- format(as.Date(CurrDate),"%Y")
	CurrMonth <- as.numeric(CurrMonthDate)
	CurrYear <- as.numeric(CurrYearDate)
	LastSix_GK <- data.frame()

# Load Last Six Months - Current Year Months
	if(CurrMonth >= 6)
	{
	  Counter = 1
	  CMon = CurrMonth - 5
	  
	  while(Counter < 7)
	  {
		LastSix_GK[Counter,1] <- c(CMon)
		LastSix_GK[Counter,2] <- c(CurrYear)
		
		Counter <- Counter + 1
		CMon <- CMon + 1
	  }
	}

# Load Last Six Months - Current and Previous Years
	if(CurrMonth < 6)
	{
	  MonthSeq <- c(8,9,10,11,12,1,2,3,4,5)
	  CMon <- grep(CurrMonth,MonthSeq)
	  Counter = 1
	  
	  while(Counter < 7)
	  {
		LastSix_GK[Counter,1] <- c(MonthSeq[CMon-5])
		
		Counter <- Counter + 1
		CMon <- CMon + 1
	  }
	  
	  Counter <- 1
	  
	  while(Counter < 7)
	  {
		if(LastSix_GK[Counter,1] > 7)
		{
		  LastSix_GK[Counter,2] <- CurrYear - 1
		}
		else
		  LastSix_GK[Counter,2] <- CurrYear
		
		Counter <- Counter + 1  
	  }
	}

# Get total tickets in each of last six months
	A <- 1

	while(A < 7)
	{
	  LastSix_GK[A,3] <- 0
	  LastSix_GK[A,4] <- 0
	  LastSix_GK[A,5] <- 0
	  LastSix_GK[A,6] <- 0
	  LastSix_GK[A,7] <- 0
	  LastSix_GK[A,8] <- 0
	  
	  A <- A + 1
	}

	C <- 1
	while(C < 7)
	{
	  LastSix_GK[C,4] <- length(GK_Parsing[(GK_Parsing$Priority == 1 & GK_Parsing$Month == LastSix_GK[C,1] & GK_Parsing$Year == LastSix_GK[C,2]),9])
	  LastSix_GK[C,6] <- length(GK_Parsing[(GK_Parsing$Priority == 2 & GK_Parsing$Month == LastSix_GK[C,1] & GK_Parsing$Year == LastSix_GK[C,2]),9])
	  LastSix_GK[C,8] <- length(GK_Parsing[(GK_Parsing$Priority == 3 & GK_Parsing$Month == LastSix_GK[C,1] & GK_Parsing$Year == LastSix_GK[C,2]),9])
	  
	  C <- C + 1
	}

# Get average outstanding for tickets - last six months
	A <- 1

	while(A < 7)
	{
	  if(is.na(sum(GK_Parsing[which(GK_Parsing$Priority == 1 & GK_Parsing$Month == LastSix_GK[A,1] & GK_Parsing$Year == LastSix_GK[A,2]),6])))
	  {
		LastSix_GK[A,3] <- 0
	  }
	  else
	  {
		LastSix_GK[A,3] <- round(sum(GK_Parsing[which(GK_Parsing$Priority == 1 & GK_Parsing$Month == LastSix_GK[A,1] & GK_Parsing$Year == LastSix_GK[A,2]),6]) / LastSix_GK[A,4],0)
	  }
	  
	  if(is.na(sum(GK_Parsing[which(GK_Parsing$Priority == 2 & GK_Parsing$Month == LastSix_GK[A,1] & GK_Parsing$Year == LastSix_GK[A,2]),6])))
	  {
		LastSix_GK[A,5] <- 0
	  }
	  else
	  {
		LastSix_GK[A,5] <- round(sum(GK_Parsing[which(GK_Parsing$Priority == 2 & GK_Parsing$Month == LastSix_GK[A,1] & GK_Parsing$Year == LastSix_GK[A,2]),6]) / LastSix_GK[A,6],0)
	  }
	  
	  if(is.na(sum(GK_Parsing[which(GK_Parsing$Priority == 3 & GK_Parsing$Month == LastSix_GK[A,1] & GK_Parsing$Year == LastSix_GK[A,2]),6])))
	  {
		LastSix_GK[A,7] <- 0
	  }
	  else
	  {
		LastSix_GK[A,7] <- round(sum(GK_Parsing[which(GK_Parsing$Priority == 3 & GK_Parsing$Month == LastSix_GK[A,1] & GK_Parsing$Year == LastSix_GK[A,2]),6]) / LastSix_GK[A,8],0)
	  }  
	  
	  A <- A + 1
	}

# Change any NaN (divide by 0) to 0
	LastSix_GK <- replace(LastSix_GK,is.na(LastSix_GK),0)

# Colate Data For Chart
	A <- 1
	df1 <- data.frame()

	while(A < 7)
	{
	  df1[A,1] <- paste(LastSix_GK[A,1]," - ",LastSix_GK[A,2])
	  
	  A <- A + 1
	}

	A <- 1

	while(A < 7)
	{
	  df1[A,2] <- LastSix_GK[A,3]
	  df1[A,3] <- LastSix_GK[A,4]
	  df1[A,4] <- LastSix_GK[A,5]
	  df1[A,5] <- LastSix_GK[A,6]
	  df1[A,6] <- LastSix_GK[A,7]
	  df1[A,7] <- LastSix_GK[A,8]
	  
	  A <- A + 1
	}

	colnames(df1) <- c("Month","Priority 1 - Avg Outstanding","Priority 1 - Tickets Created","Priority 2 - Avg Outstanding","Priority 2 - Tickets Created","Priority 3 - Avg Outstanding","Priority 3 - Tickets Created")

### Create Plot Data - Victor Mine
	CurrDate <- Sys.Date()
	CurrMonthDate <- format(as.Date(CurrDate),"%m")
	CurrYearDate <- format(as.Date(CurrDate),"%Y")
	CurrMonth <- as.numeric(CurrMonthDate)
	CurrYear <- as.numeric(CurrYearDate)
	LastSix_VM <- data.frame()

# Load Last Six Months - Current Year Months
	if(CurrMonth >= 6)
	{
	  Counter = 1
	  CMon = CurrMonth - 5
	  
	  while(Counter < 7)
	  {
		LastSix_VM[Counter,1] <- c(CMon)
		LastSix_VM[Counter,2] <- c(CurrYear)
		
		Counter <- Counter + 1
		CMon <- CMon + 1
	  }
	}

# Load Last Six Months - Current and Previous Years
	if(CurrMonth < 6)
	{
	  MonthSeq <- c(08,09,10,11,12,01,02,03,04,05)
	  CMon <- grep(CurrMonth,MonthSeq)
	  Counter = 1
	  
	  while(Counter < 7)
	  {
		LastSix_VM[Counter,1] <- c(MonthSeq[CMon-5])
		
		Counter <- Counter + 1
		CMon <- CMon + 1
	  }
	  
	  Counter <- 1
	  
	  while(Counter < 7)
	  {
		if(LastSix_VM[Counter,1] > 7)
		{
		  LastSix_VM[Counter,2] <- CurrYear - 1
		}
		else
		  LastSix_VM[Counter,2] <- CurrYear
		
		Counter <- Counter + 1  
	  }
	}

# Get total tickets in each of last six months
	A <- 1

	while(A < 7)
	{
	  LastSix_VM[A,3] <- 0
	  LastSix_VM[A,4] <- 0
	  LastSix_VM[A,5] <- 0
	  LastSix_VM[A,6] <- 0
	  LastSix_VM[A,7] <- 0
	  LastSix_VM[A,8] <- 0
	  
	  A <- A + 1
	}

	C <- 1

	while(C < 7)
	{
	  LastSix_VM[C,4] <- length(VM_Parsing[(VM_Parsing$Priority == 1 & as.numeric(VM_Parsing$Month == LastSix_VM[C,1]) & VM_Parsing$Year == LastSix_VM[C,2]),9])
	  LastSix_VM[C,6] <- length(VM_Parsing[(VM_Parsing$Priority == 2 & as.numeric(VM_Parsing$Month == LastSix_VM[C,1]) & VM_Parsing$Year == LastSix_VM[C,2]),9])
	  LastSix_VM[C,8] <- length(VM_Parsing[(VM_Parsing$Priority == 3 & as.numeric(VM_Parsing$Month == LastSix_VM[C,1]) & VM_Parsing$Year == LastSix_VM[C,2]),9])
	  
	  C <- C + 1
	}

# Get average outstanding for tickets - last six months
	A <- 1

	while(A < 7)
	{
	  if(is.na(sum(VM_Parsing[which(VM_Parsing$Priority == 1 & VM_Parsing$Month == LastSix_VM[A,1] & VM_Parsing$Year == LastSix_VM[A,2]),6])))
	  {
		LastSix_VM[A,3] <- 0
	  }
	  else
	  {
		LastSix_VM[A,3] <- round(sum(VM_Parsing[which(VM_Parsing$Priority == 1 & VM_Parsing$Month == LastSix_VM[A,1] & VM_Parsing$Year == LastSix_VM[A,2]),6]) / LastSix_VM[A,4],0)
	  }
	  
	  if(is.na(sum(VM_Parsing[which(VM_Parsing$Priority == 2 & VM_Parsing$Month == LastSix_VM[A,1] & VM_Parsing$Year == LastSix_VM[A,2]),6])))
	  {
		LastSix_VM[A,5] <- 0
	  }
	  else
	  {
		LastSix_VM[A,5] <- round(sum(VM_Parsing[which(VM_Parsing$Priority == 2 & VM_Parsing$Month == LastSix_VM[A,1] & VM_Parsing$Year == LastSix_VM[A,2]),6]) / LastSix_VM[A,6],0)
	  }
	  
	  if(is.na(sum(VM_Parsing[which(VM_Parsing$Priority == 3 & VM_Parsing$Month == LastSix_VM[A,1] & VM_Parsing$Year == LastSix_VM[A,2]),6])))
	  {
		LastSix_VM[A,7] <- 0
	  }
	  else
	  {
		LastSix_VM[A,7] <- round(sum(VM_Parsing[which(VM_Parsing$Priority == 3 & VM_Parsing$Month == LastSix_VM[A,1] & VM_Parsing$Year == LastSix_VM[A,2]),6]) / LastSix_VM[A,8],0)
	  }  
	  
	  A <- A + 1
	}

# Change any NaN (divide by 0) to 0
	LastSix_VM <- replace(LastSix_VM,is.na(LastSix_VM),0)

# Colate Data For Chart
	A <- 1
	df2 <- data.frame()

	while(A < 7)
	{
	  df2[A,1] <- paste(LastSix_VM[A,1]," - ",LastSix_VM[A,2])
	  
	  A <- A + 1
	}

	A <- 1

	while(A < 7)
	{
	  df2[A,2] <- LastSix_VM[A,3]
	  df2[A,3] <- LastSix_VM[A,4]
	  df2[A,4] <- LastSix_VM[A,5]
	  df2[A,5] <- LastSix_VM[A,6]
	  df2[A,6] <- LastSix_VM[A,7]
	  df2[A,7] <- LastSix_VM[A,8]
	  
	  A <- A + 1
	}

	colnames(df2) <- c("Month","Priority 1 - Avg Outstanding","Priority 1 - Tickets Created","Priority 2 - Avg Outstanding","Priority 2 - Tickets Created","Priority 3 - Avg Outstanding","Priority 3 - Tickets Created")

### Plotting All
	library(reshape2)
	library(ggplot2)

# Y Scale - Find Highest # in each of GK and VM Frames
	DF_GK_Scale <- max(df1[2:7], na.rm=TRUE)
	DF_VM_Scale <- max(df2[2:7], na.rm=TRUE)

# Round up to nearest 10
	DF_GK_Scale <- ceiling(DF_GK_Scale/10)*10
	DF_VM_Scale <- ceiling(DF_VM_Scale/10)*10

# Melt data to create long vs wide / Plot - GK
	df_G <- melt(df1,id.vars = "Month")
	
	MonOrder <- c(df1[1,1],df1[2,1],df1[3,1],df1[4,1],df1[5,1],df1[6,1])
	cbPalette <- c("#000000", "#E69F00", "#56B4E9", "#009E73", "#F0E442", "#0072B2", "#D55E00", "#CC79A7")

	ggplot(df_G, aes(x=Month, y=value, fill=variable)) + 
	scale_x_discrete(limits = MonOrder) +  
	geom_bar(stat='identity', position ='dodge', colour = 'black') +
	xlab("Month/Year") + ylab("New Notifications/Average Outstanding") +
	ggtitle("Gahcho Kue Mine - SAP Notifications - Last 6 Months") +
	scale_fill_manual(values=cbPalette) +
	labs(fill = "Legend") +
	scale_y_continuous(limits = c(0,DF_GK_Scale), breaks = seq(0,DF_GK_Scale,10), minor_breaks = seq(10,110,20))

# Melt data to create long vs wide / Plot - VM
	df_M <- melt(df2,id.vars = "Month")
	
	MonOrder <- c(df2[1,1],df2[2,1],df2[3,1],df2[4,1],df2[5,1],df2[6,1])
	cbPalette <- c("#000000", "#E69F00", "#56B4E9", "#009E73", "#F0E442", "#0072B2", "#D55E00", "#CC79A7")

	ggplot(df_M, aes(x=Month, y=value, fill=variable)) + 
	scale_x_discrete(limits = MonOrder) +  
	geom_bar(stat='identity', position ='dodge', colour = 'black') +
	xlab("Month/Year") + ylab("New Notifications/Average Outstanding") +
	ggtitle("Victor Mine - SAP Notifications - Last 6 Months") +
	scale_fill_manual(values=cbPalette) +
	labs(fill = "Legend") +
	scale_y_continuous(limits = c(0,DF_VM_Scale), breaks = seq(0,DF_VM_Scale,10), minor_breaks = seq(10,115,20))

	# Set Working Directory
	setwd("C:/Users/ken.bartlett/Desktop/My Stuff/R/Weekly_Reports")
	