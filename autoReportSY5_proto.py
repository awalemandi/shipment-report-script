import pandas as pd
import datetime as dt
import re

# Creating a date variable for comparison in Eastern format
today = dt.date.today()
yesterday = dt.date.today() - dt.timedelta(days=1)
weekAgo = dt.date.today() - dt.timedelta(days=7)


def autoReport(shipmentRegister) :

    # creating dataframes from the excel file
    dfShipments = pd.read_excel(shipmentRegister, sheet_name='ShipmentRegister')
    dfHandOverLog = pd.read_excel(shipmentRegister, sheet_name='HandoverLog')
    dfHandOverLog[['Handover'], ['Date\n(DD-MMM-YY)']] = pd.to_datetime(dfHandOverLog[['Handover'], ['Date\n(DD-MMM-YY)']])

    # filter the dataframe by getting the rows with shipment IDs not ending with a - and digit (just the original ticket IDs)
    dfTicketsForShipmentRegister = dfShipments.loc[(dfShipments['Shipment ID\n(SO#, Temporary Shipment ID, Storage Event ID)'].str.contains('.-\d$') == False) & (dfShipments['Status'] == 'Inside The Secure Store')]
    dfTicketsForHandOver = dfHandOverLog.loc[dfHandOverLog['Shipment ID\n(SO#)'].str.contains('.-\d$') == False]

    # filter the number of tickets inside SS over x days
    ticketsInStore = dfTicketsForShipmentRegister.shape[0]
    ticketsOver21 = dfTicketsForShipmentRegister.loc[(dfTicketsForShipmentRegister['Status'] == 'Inside The Secure Store') & (dfTicketsForShipmentRegister['Age'] > 21)].shape[0]
    ticketsOver14 = dfTicketsForShipmentRegister.loc[(dfShipments['Status'] == 'Inside The Secure Store') & (dfTicketsForShipmentRegister['Age'].between(14, 22))].shape[0]
    ticketsOver7 = dfTicketsForShipmentRegister.loc[(dfShipments['Status'] == 'Inside The Secure Store') & (dfTicketsForShipmentRegister['Age'].between(6, 15))].shape[0]

    # filter number of shipments delivered/picked up in a week
    shipmentDeliveriesThisWeek = dfTicketsForShipmentRegister.loc[dfTicketsForShipmentRegister['Age'] <= 7].shape[0]
    shipmentCollectionsThisWeek = dfTicketsForHandOver.loc[(dfTicketsForHandOver['Handover'], ['Date\n(DD-MMM-YY)'] <= weekAgo)].shape[0]

    # filter the number of items inside SS over x days
    totalShipmentInside = dfShipments.loc[dfShipments['Status'] == 'Inside The Secure Store'].shape[0]
    itemsOver21 = dfShipments.loc[(dfShipments['Status'] == 'Inside The Secure Store') & (dfShipments['Age'] > 21)].shape[0]
    itemsOver14 = dfShipments.loc[(dfShipments['Status'] == 'Inside The Secure Store') & (dfShipments['Age'].between(14, 22))].shape[0]
    itemsOver7 = dfShipments.loc[(dfShipments['Status'] == 'Inside The Secure Store') & (dfShipments['Age'].between(6, 15))].shape[0]

    # filter number of items delivered/picked up in a week
    deliveriesThisWeek = dfShipments.loc[dfShipments['Age'] <= 7].shape[0]
    collectionsThisWeek = dfHandOverLog.loc[(dfHandOverLog['Handover'], ['Date\n(DD-MMM-YY)'] < weekAgo)].shape[0]

    #print out outputs
    print("Welcome to Auto Report for the week: " + str(weekAgo) + " to " + str(yesterday))
    print('')
    # print('Items:')

    print('Total number of outstanding items over 21 days (21+): ' + str(itemsOver21) + ' items')
    print('Total number of outstanding items over 14 days (15-21): ' + str(itemsOver14) + ' items')
    print('Total number of outstanding items over 7 days (7-14): ' + str(itemsOver7) + ' items')
    print('Total number of physical shipments inside secure store: ' + str(totalShipmentInside) + ' items')
    print('')
    print('Total number of deliveries this week: ' + str(deliveriesThisWeek) + ' items')
    print('Total number of collections this week: ' + str(collectionsThisWeek) + ' items')

    print('')
    print('Tickets:')
    print('Total number of outstanding shipments over 21 days (21+): ' + str(ticketsOver21) + ' shipments')
    print('Total number of outstanding shipments over 14 days (15-21): ' + str(ticketsOver14) + ' shipments')
    print('Total number of outstanding shipments over 7 days (7-14): ' + str(ticketsOver7) + ' shipments')
    print('Total number of shipments inside secure store: ' + str(ticketsInStore) + ' shipments')
    print('')
    print('Total number of shipment deliveries this week: ' + str(shipmentDeliveriesThisWeek) + ' shipments')
    print('Total number of shipment collections this week: ' + str(shipmentCollectionsThisWeek) + ' shipments')
    input()

autoReport('SY5-Shipment_Register_MasterCopy.xlsx')


