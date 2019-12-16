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
perform  = pd.read_csv('performance_by_SKU_result.csv')
new      = pd.read_csv('future_inventory_parameters_result.csv')
his      = pd.read_csv('historical_profit_by_SKU_result.csv')
storage  = pd.read_csv('R2_StorageCost.csv')
bo       = pd.read_csv('bonus.csv')
init     = pd.read_csv('R2_Initial_Inventories.csv')

#create output file
df = pd.DataFrame(prod, columns      = ['ProductId'])
df['EstimatedSalesQuantityKg']       = 0.0
df['EstimatedSalesValue']            = 0.0
df['EstimatedPurchaseCost']	         = 0.0
df['EstimatedInboundTransportCost']  = 0.0
df['EstimatedOutboundTransportCost'] = 0.0
df['EstimatedStorageCost']           = 0.0
df['EstimatedProfit']                = 0.0
df['HistoricalProfit']               = 0.0
df['ProfitDifference']               = 0.0
																
#the loop
for i in range(len(df['ProductId'])):

	#Cost - base on "R2_SOLine"
	cost_so  = (so.loc[so['SoProductId'] == df['ProductId'][i], 'SoUnitCost'].mean())
	#Cost - base on "R2_POLine"
	cost_po  = (po.loc[po['PoProductId'] == df['ProductId'][i], 'PoUnitCost'].mean())
	#Cost - inbound
	a        = ro.loc[ro['RoProductId'] == df['ProductId'][i], 'RoSupplierId'].min()
	cost_in  = inbound.loc[inbound['SupplierId'] == a, 'TransportUnitCost'].mean()
	#Cost - outbound
	b        = to.loc[to['ToProductId'] == df['ProductId'][i], 'ToCustomerId'].min()
	cost_out = outbound.loc[outbound['CustomerId'] == b, 'TransportUnitCost'].mean()
	#Cost storage
	store_cost = storage.loc[storage['ProductId'] == df['ProductId'][i], 'UnitCost'].mean()

	#Caculate EstimatedSalesQuantityKg
	df['EstimatedSalesQuantityKg'][i]    = perform['TotalOrderQuantityKg'][i] * new['NewServiceLevelPercent'][i]

	#Caculate EstimatedSalesValue
	df['EstimatedSalesValue'][i]         = df['EstimatedSalesQuantityKg'][i] * cost_so


	################################
	
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
	################################	


	#Caculate EstimatedPurchaseCost
	df['EstimatedPurchaseCost'][i]       = perform['TotalOrderQuantityKg'][i] + new['NewSafetyStockKg'][i] - openingInventory - bo['XXX'][i]



	#Caculate EstimatedInboundTransportCost
	df['EstimatedInboundTransportCost'][i]  = df['EstimatedSalesQuantityKg'][i] * cost_in

	#Caculate EstimatedOutboundTransportCost
	df['EstimatedOutboundTransportCost'][i] = df['EstimatedSalesQuantityKg'][i] * cost_out
	
	#Caculate EstimatedStorageCost
	start            = datetime.strptime(so.SoEnteredDate.min(), '%Y-%m-%d').date()
	end              = datetime.strptime(so.SoEnteredDate.max(), '%Y-%m-%d').date()
	investigatedDays = (end - start).days + 1
	df['EstimatedStorageCost'] = new['NewAverageInventoryKg'] * investigatedDays * store_cost
		
	#Caculate EstimatedProfit
	#Caculate EstimatedProfit
	df['EstimatedProfit'][i] = df['EstimatedSalesValue'][i] - df['EstimatedPurchaseCost'][i] - df['EstimatedInboundTransportCost'][i] - df['EstimatedOutboundTransportCost'][i] - df['EstimatedStorageCost'][i] 
	
	#Caculate HistoricalProfit
	df['HistoricalProfit'][i]   = his['Profit'][i]
	#Caculate ProfitDifference
	df['ProfitDifference'][i]   = df['EstimatedProfit'][i] - df['HistoricalProfit'][i] 

#check
print(df)
print(investigatedDays)

#export
df.to_csv(r'profit_comparison_result.csv', float_format = '%.3f', index = False, decimal='.')