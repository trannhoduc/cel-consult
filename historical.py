import pandas as pd
from pandas import ExcelWriter, ExcelFile
from datetime import datetime
from datetime import timedelta

# open input files
prod     = pd.read_csv('R2_ProductMaster.csv')
ro       = pd.read_csv('R2_ROLine.csv')
so       = pd.read_csv('R2_SOLine.csv')
po       = pd.read_csv('R2_POLine.csv')
to       = pd.read_csv('R2_TOLine.csv')
inbound  = pd.read_csv('R2_InboundTransportRate.csv')
outbound = pd.read_csv('R2_OutboundTransportRate.csv')
storage  = pd.read_csv('R2_StorageCost.csv')
init     = pd.read_csv('R2_Initial_Inventories.csv')
perform  = pd.read_csv('performance_by_SKU_result.csv')

#create output file
df = pd.DataFrame(prod, columns = ['ProductId'])
df['PurchaseCost']              = 0.0
df['InboundTransportCost']      = 0.0
df['OutboundTransportCost']	    = 0.0
df['StorageCost']               = 0.0
df['Profit']                    = 0.0

#the loop
for i in range(len(df['ProductId'])):
	store_cost = storage.loc[storage['ProductId'] == df['ProductId'][i], 'UnitCost']

	#Caculate purchase cost
	df['PurchaseCost'][i]           = (
		ro.loc[ro['RoProductId']    == df['ProductId'][i], 
		'RoInKg']).sum() * (po.loc[po['PoProductId'] == df['ProductId'][i], 'PoUnitCost'].mean())

	#Caculate Inbound Cost
	a    = ro.loc[ro['RoProductId'] == df['ProductId'][i], 'RoSupplierId'].min()
	#Therefore
	cost = inbound.loc[inbound['SupplierId'] == a, 'TransportUnitCost'].mean()
	df['InboundTransportCost'][i]    =  ro.loc[ro['RoProductId']    == df['ProductId'][i], 'RoInKg'].sum() * cost

	#Caculate Outbound Cost
	a    = to.loc[to['ToProductId'] == df['ProductId'][i], 'ToCustomerId'].min()
	#Therefore
	cost = outbound.loc[outbound['CustomerId'] == a, 'TransportUnitCost'].mean()
	df['OutboundTransportCost'][i]    =  to.loc[to['ToProductId']    == df['ProductId'][i], 'ToInKg'].sum() * cost

	#Caculate Storage Profit
	start            = datetime.strptime(so.SoEnteredDate.unique().min(), '%Y-%m-%d').date()
	end              = datetime.strptime(so.SoEnteredDate.unique().max(), '%Y-%m-%d').date()
	periodDays       = (end - start).days
	init_value       = init.loc[init['ProductId'] == df['ProductId'][i], 'Quantity'].mean()
	storageCost      = 0.0
	investigatedDays = to['ToDepartureDate'].nunique()
	#
	for j in range(periodDays):
		if j == 0:
			openingInventory = init_value
			closing_value    = openingInventory
		else:
			day = start + timedelta(days = j - 1)
			closing_value = openingInventory

			inPut = ro.loc[
				(ro['RoProductId'] == df['ProductId'][i]) & (ro['RoArrivalDate'] == str(day)), 
					'RoInKg'].sum()
			outPut = to.loc[
				(to['ToProductId'] == df['ProductId'][i]) & (to['ToDepartureDate'] == str(day)), 
					'ToInKg'].sum()

			openingInventory = closing_value + inPut - outPut

		storageCost      = storageCost + openingInventory*store_cost
	#
	df['StorageCost'][i] = storageCost

	#Profit
	df['Profit'][i] = perform['SalesValue'][i] - df['PurchaseCost'][i] - df['InboundTransportCost'][i] - df['OutboundTransportCost'][i] - df['StorageCost'][i]

#check
print(df)

#export
df.to_csv(r'historical_profit_by_SKU_result.csv', float_format = '%.3f', index = False)