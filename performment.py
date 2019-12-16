import pandas as pd
from pandas import ExcelWriter, ExcelFile
from datetime import datetime
from datetime import timedelta

# open input files
prod     = pd.read_csv('R2_ProductMaster.csv')
ro       = pd.read_csv('R2_ROLine.csv')
po       = pd.read_csv('R2_POLine.csv')
to       = pd.read_csv('R2_TOLine.csv')
so       = pd.read_csv('R2_SOLine.csv')
inbound  = pd.read_csv('R2_InboundTransportRate.csv')
outbound = pd.read_csv('R2_OutboundTransportRate.csv')
init     = pd.read_csv('R2_Initial_Inventories.csv')

#create output file
df = pd.DataFrame(prod, columns = ['ProductId'])
df['TotalOrderQuantityKg']              = 0.0
df['TotalShippedQuantityKg']            = 0.0
df['DailyDemandQuantityKg']	            = 0.0
df['DailyDemandValue']                  = 0.0
df['SalesValue']                        = 0.0
df['CumulatedShareInSalesValuePercent'] = 0.0
df['AbcClass']                          = 0.0
df['AverageInventoryKg']                = 0.0
df['DioDay']                            = 0.0
df['ServiceLevelPercent']               = 0.0
									
#the loop
for i in range(len(df['ProductId'])):

	#Cost - base on "R2_SOLine"
	cost_so = (so.loc[so['SoProductId'] == df['ProductId'][i], 'SoUnitCost'].mean())

	#Caculate TotalOrderQuantityKg
	df['TotalOrderQuantityKg'][i]       = (
		so.loc[so['SoProductId']        == df['ProductId'][i], 
		'SoQuantity']).sum() 

	#Caculate TotalShipQuantityKg
	df['TotalShippedQuantityKg'][i]     = (
		to.loc[to['ToProductId']        == df['ProductId'][i], 
		'ToQuantity']).sum() 

	#Caculate DailyDemandQuantityKg
	numOfDays = (datetime.strptime(so.SoEnteredDate.unique().max(), '%Y-%m-%d').date() 
				- datetime.strptime(so.SoEnteredDate.unique().min(), '%Y-%m-%d').date()).days
	df['DailyDemandQuantityKg'][i]      = (
		so.loc[so['SoProductId']        == df['ProductId'][i], 
		'SoQuantity']).sum()  / (numOfDays + 1)

	#Caculate DailyDemandValue
	df['DailyDemandValue'][i]           = (
		so.loc[so['SoProductId']        == df['ProductId'][i], 
		'SoQuantity']).sum() * cost_so / (numOfDays + 1)

	#Caculate SalesValue
	df['SalesValue'][i]                 = (
		to.loc[to['ToProductId']        == df['ProductId'][i], 
		'ToQuantity']).sum() * cost_so

	#Caculate AverageInventoryKg
	#
	#Caculate Sum of Opening Inventory
	#
	start            = datetime.strptime(so.SoEnteredDate.min(), '%Y-%m-%d').date()
	end              = datetime.strptime(so.SoEnteredDate.max(), '%Y-%m-%d').date()
	periodDays       = (end - start).days
	init_value       = init.loc[init['ProductId'] == df['ProductId'][i], 'Quantity'].mean()
	sumOpening       = 0.0
	investigatedDays = periodDays + 1
	#
	for j in range(periodDays):
		if j == 0:
			openingInventory = init_value
			closing_value    = openingInventory
		else:
			day = start + timedelta(days = j - 1)
			inPut = ro.loc[
				(ro['RoProductId'] == df['ProductId'][i]) & (ro['RoArrivalDate'] == str(day)), 
					'RoInKg'].sum()
			outPut = to.loc[
				(to['ToProductId'] == df['ProductId'][i]) & (to['ToDepartureDate'] == str(day)), 
					'ToInKg'].sum()

			openingInventory = closing_value + inPut - outPut
			closing_value    = openingInventory

		sumOpening       = sumOpening + openingInventory
	#
	df['AverageInventoryKg'][i] = sumOpening / investigatedDays

	#Caculate DioDay   
	df['DioDay'][i]        = df['AverageInventoryKg'][i] / df['DailyDemandQuantityKg'][i]

	#Caculate ServiceLevelPercent
	a  = df['TotalShippedQuantityKg'][i] / df['TotalOrderQuantityKg'][i]
	df['ServiceLevelPercent'][i] = a * 100.0


# for i in range(len(df['ProductId'])):
# 	#Caculate CumulatedShareInSalesValuePercent
# 	b  = df['SalesValue'][i] / df['SalesValue'].sum() * 100.0
# 	if i == 0:
# 		df['CumulatedShareInSalesValuePercent'][i] = b
# 	else:
# 		df['CumulatedShareInSalesValuePercent'][i] = df['CumulatedShareInSalesValuePercent'][i - 1] + b
	
# 	#Determine AbcClass
# 	if df['CumulatedShareInSalesValuePercent'][i] < 0.8 :
# 		df['AbcClass'][i] =  "A" 
# 	elif df['CumulatedShareInSalesValuePercent'][i] > 0.95 :
# 		df['AbcClass'][i] =  "C" 
# 	else:
# 		df['AbcClass'][i] =  "B" 

	
#check
print(df)

#export
df.to_csv(r'performance_by_SKU_result.csv', float_format = '%.3f', index = False, decimal='.')