									
import pandas as pd
from pandas import ExcelWriter, ExcelFile
from datetime import datetime
import math
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
service  = pd.read_csv('R2_Projected_SL_By_Segmentation.csv')
coverage = pd.read_csv('R2_Projected_CoverageDay_By_Class.csv')

#create output file
df = pd.DataFrame(prod, columns      = ['ProductId'])
df['AverageLeadTimeDay']             = 0.0
df['StandardDeviationOfLeadTimeDay'] = 0.0
df['StandardDeviationOfDemandKg']	 = 0.0
df['AbcClass']                       = 0.0
df['NormalizedClass']                = 0.0
df['NewServiceLevelPercent']         = 0.0
df['NewServiceLevelZScore']          = 0.0
df['NewCoveragePeriodDay']           = 0.0
df['NewSafetyStockKg']               = 0.0
df['NewAverageInventoryKg']          = 0.0
																
#the loop
for i in range(len(df['ProductId'])):

	#Cost - base on "R2_SOLine"
	cost_so = (so.loc[so['SoProductId'] == df['ProductId'][i], 'SoUnitCost'].mean())
	#Cost - base on "R2_POLine"
	cost_po = (po.loc[po['PoProductId'] == df['ProductId'][i], 'PoUnitCost'].mean())

	#Caculate AverageLeadTimeDay
	po['PoEnteredDate'] = pd.to_datetime(po['PoEnteredDate'], format='%Y-%m-%d')
	ro['RoArrivalDate'] = pd.to_datetime(ro['RoArrivalDate'], format='%Y-%m-%d')
	averageSupplyTime   = 0.0
	times = 0.0

	# ROday = ro.loc[ro['RoProductId'] == df['ProductId'][i], 'PoOrderLineId']
	for m in range(len(ro['RoProductId'])):
		if ro['RoProductId'][m] == df['ProductId'][i]:
			RO_Arrival = ro['RoArrivalDate'][m]
			PO_Enter   = po.loc[po['PoOrderLineId'] == ro['PoOrderLineId'][m], 'PoEnteredDate'].min()
			supplyTime = (RO_Arrival - PO_Enter).days
			averageSupplyTime = averageSupplyTime + supplyTime
			times = times + 1

	averageSupplyTime = averageSupplyTime / times
	df['AverageLeadTimeDay'][i] = averageSupplyTime

	#Caculate StandardDeviationOfLeadTimeDay
	tu_so = 0.0
	for m in range(len(ro['RoProductId'])):
		if ro['RoProductId'][m] == df['ProductId'][i]:
			RO_Arrival = ro['RoArrivalDate'][m]
			PO_Enter   = po.loc[po['PoOrderLineId'] == ro['PoOrderLineId'][m], 'PoEnteredDate'].min()
			supplyTime = (RO_Arrival - PO_Enter).days
			tu_so = tu_so + (supplyTime - df['AverageLeadTimeDay'][i]) ** 2

	df['StandardDeviationOfLeadTimeDay'][i] = math.sqrt(tu_so / times)

	#Caculate StandardDeviationOfDemandKg
	numOfDays = (datetime.strptime(so.SoEnteredDate.unique().max(), '%Y-%m-%d').date() 
				- datetime.strptime(so.SoEnteredDate.unique().min(), '%Y-%m-%d').date()).days
	average   = (
		so.loc[so['SoProductId']        == df['ProductId'][i], 
		'SoQuantity']).sum()  / (numOfDays + 1)
	n         = so.SoEnteredDate.nunique()

	start            = datetime.strptime(so.SoEnteredDate.min(), '%Y-%m-%d').date()
	end              = datetime.strptime(so.SoEnteredDate.max(), '%Y-%m-%d').date()
	times            = 0.0
	tu               = 0.0
	#
	for j in range(numOfDays):
		day   = start + timedelta(days = j)
		daily = so.loc[
			(so['SoProductId'] == df['ProductId'][i]) & (so['SoEnteredDate'] == str(day)), 
				'SoQuantity'].sum()
		tu    =  tu + (daily - average)**2

	df['StandardDeviationOfDemandKg'][i] = math.sqrt(tu / 153.0)

	#Caculate abcClass
	df['AbcClass'][i]  =  perform['AbcClass'][i]

	#Caculate NormalizedClass
	#
	NormalizedClass = df['StandardDeviationOfDemandKg'][i] / average 
	if NormalizedClass < 1:
		df['NormalizedClass'][i] = "L"
	elif NormalizedClass > 2:
		df['NormalizedClass'][i] = "H"
	else:
		df['NormalizedClass'][i] = "M"

	#Caculate NewServiceLevelPercent
	new1 = service.loc[(service['ClassAbc'] == df['AbcClass'][i]) & (service['VarianceLevel'] == df['NormalizedClass'][i]), 
		'ProjectedSl'].mean()
	df['NewServiceLevelPercent'][i] = new1

	# Caculate NewCoveragePeriodDay
	# 
	new2 = coverage.loc[coverage['ClassAbc'] == df['AbcClass'][i], 
		'CoverageByDay'].mean()
	df['NewCoveragePeriodDay'][i] = new2


	#Caculate NewServiceLevelZScore
	#
	new3 = service.loc[(service['ClassAbc'] == df['AbcClass'][i]) & (service['VarianceLevel'] == df['NormalizedClass'][i]), 
		'Z'].mean()
	df['NewServiceLevelZScore'][i] = new3
	#
	#Caculate NewSafetyStockKg
	#
	df['NewSafetyStockKg'][i] = df['NewServiceLevelPercent'][i] * math.sqrt(df['AverageLeadTimeDay'][i]*(df['StandardDeviationOfDemandKg'][i])**2 +
			(average*df['StandardDeviationOfLeadTimeDay'][i])**2)

	#####
	df['NewAverageInventoryKg'][i] = df['NewSafetyStockKg'][i] + (df['NewCoveragePeriodDay'][i]/2) * average

#check
print(df)

#export
df.to_csv(r'future_inventory_parameters_result.csv', float_format = '%.3f', index = False, decimal='.')