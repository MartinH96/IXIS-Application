# The following functions take the data from the two excel documents and
# produce an Excel document with further information in them

install.packages("openxlsx")
library("openxlsx")

setwd("C:/Users/humph/OneDrive/Documents/Job Search/IXIS Documents")

session_counts = read.csv("DataAnalyst_Ecom_data_sessionCounts.csv",
				  header = TRUE)
attach(session_counts)

# dim_date needs to be numeric, not a string
format_date = as.Date(dim_date, format = "%m/%d/%y")
date_month = format(format_date, "%y/%m")

# Aggregates the Session Counts table by Month and device
agg_session = setNames(
			     aggregate(
					   subset(session_counts,
						    select = -c(dim_browser,
								    dim_deviceCategory,
								    dim_date)
						    ),
					   by = list(date_month,
							 dim_deviceCategory),
					   FUN = sum),
			     c("month", "device", "sessions", "transactions",
				 "qty")
				 )

ecr = c(agg_session$transactions / agg_session$sessions)

# Add ECR to the aggregated table
agg_session_1 = agg_session
agg_session_1$ecr = ecr

adds_to_cart = read.csv("DataAnalyst_Ecom_data_addsToCart.csv",
				header = T)
attach(adds_to_cart)

# Aggregates the Session Counts table by month
agg_month = setNames(
			   aggregate(
					 subset(session_counts,
						  select = -c(dim_browser,
								  dim_deviceCategory,
								  dim_date)
						  ),
					 by = list(date_month),
					 FUN = sum),
			   c("month", "sessions", "transactions", "qty")
			   )

month_ecr = c(agg_month$transactions / agg_month$sessions)

# Adds ECR as an extra metric
agg_month_1 = agg_month
agg_month_1$month_ecr = month_ecr

month_val = agg_month_1
month_val$addsToCart = addsToCart

recent_month = tail(month_val, n = 2)

# Code is more complex here in case extra months are added
abs_diff = tail(recent_month[, -1], -1) - head(recent_month[, -1], -1)

rel_diff = (abs_diff / head(recent_month[, -1], -1)) *  100

diffs = rbind(abs_diff, rel_diff)
rownames(diffs) = c("abs_diff", "rel_diff")

rownames(recent_month) = c("May_13", "June_13")
options(scipen = 999)
options(digits = 4)

month_comp = rbind(recent_month[, -1], diffs)

# Add extra column so names will remain when generated in Excel
month_comp_1 = data.frame(month = row.names(month_comp), month_comp)

sheets = list("session_counts" = agg_session_1, "month_comp" = month_comp_1)

# Default table color is blue, but green looks nicer
style = createStyle(fontColour = "#FFFFFF",
			  fgFill = "#9BBB59",
			  valign = "center",
			  textDecoration = "Bold",
			  border = "TopBottomLeftRight"
			  )
ixis = write.xlsx(sheets,
			file = "agg_comp.xlsx",
			asTable = TRUE,
			row.names = TRUE,
			col.names = TRUE,
			headerStyle = style
			)