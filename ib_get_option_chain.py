"""
This script will access the IB API and download to excel the option chain for the underlying
entered in the excel file
"""

#===============================================================================
# LIBRARIES
#===============================================================================
from ib.opt import ibConnection, message
from ib.ext.Contract import Contract
import quantacademy.excel_management as excel
import pandas as pd
from time import sleep
from collections import defaultdict

#===============================================================================
# Class IB_API
#===============================================================================
class IB_API:
    """
    This class will establish a connection to IB and group the different 
    operations
    """
    
    # Variables
    d_ticker_reqId = {}
    reqId = 1
    d_opt_contracts = defaultdict(dict)
    d_contracts = {}
    
    # Functions
    def __init__(self):
        """
        Connection to the IB API
        """
        print "Calling connection"
        # Creation of Connection class
        self.connection = ibConnection()
        # Register data handlers
        self.connection.registerAll(self.process_messages)
        # Connect
        self.connection.connect()
        
    def process_messages(self, msg):
        """
        Function that indicates how to process each different message
        """
        if msg.typeName == "contractDetails":
            print msg
            opt_contract = msg.values()[1]
            self.save_option_contracts_to_dict(opt_contract)
        elif msg.typeName == "tickPrice":
            field = msg.values()[1]
            price = msg.values()[2]
            localSymbol = self.d_ticker_reqId[msg.values()[0]].m_localSymbol
            if field == 1:
                self.d_opt_contracts[localSymbol]['bid'] = str(price)
            elif field == 2:
                self.d_opt_contracts[localSymbol]['ask'] = str(price)
            elif field == 9:
                self.d_opt_contracts[localSymbol]['close'] = str(price)
        elif msg.typeName == "tickOptionComputation":
            localSymbol = self.d_ticker_reqId[msg.values()[0]].m_localSymbol
            field = msg.values()[1]
            price = msg.values()[2]
            delta = msg.values()[3]
            impliedVolatility = msg.values()[4]
            optPrice = msg.values()[5]
            pvDividend = msg.values()[6]
            gamma = msg.values()[7]
            vega = msg.values()[8]
            theta = msg.values()[9]
            undPrice = None
            if field == 10:
                self.d_opt_contracts[localSymbol]['bid_delta'] = str(delta)
                self.d_opt_contracts[localSymbol]['bid_impliedVolatility'] = str(impliedVolatility)
                self.d_opt_contracts[localSymbol]['bid_optPrice'] = str(optPrice)
                self.d_opt_contracts[localSymbol]['bid_pvDividend'] = str(pvDividend)
                self.d_opt_contracts[localSymbol]['bid_gamma'] = str(gamma)
                self.d_opt_contracts[localSymbol]['bid_vega'] = str(vega)
                self.d_opt_contracts[localSymbol]['bid_theta'] = str(theta)
                self.d_opt_contracts[localSymbol]['bid_undPrice'] = str(undPrice)
            elif field == 11:
                self.d_opt_contracts[localSymbol]['ask_delta'] = str(delta)
                self.d_opt_contracts[localSymbol]['ask_impliedVolatility'] = str(impliedVolatility)
                self.d_opt_contracts[localSymbol]['ask_optPrice'] = str(optPrice)
                self.d_opt_contracts[localSymbol]['ask_pvDividend'] = str(pvDividend)
                self.d_opt_contracts[localSymbol]['ask_gamma'] = str(gamma)
                self.d_opt_contracts[localSymbol]['ask_vega'] = str(vega)
                self.d_opt_contracts[localSymbol]['ask_theta'] = str(theta)
                self.d_opt_contracts[localSymbol]['ask_undPrice'] = str(undPrice)
            elif field == 24:
                self.d_opt_contracts[localSymbol]['iv'] = str(price)
                
        else:
            print msg
            
    def get_contract_details(self, reqId, contract_values):
        """
        Call for all the options contract for the underlying
        """
        print "Calling Contract Details"
        # Contract creation
        contract = Contract()
        contract.m_symbol = contract_values['m_symbol']
        contract.m_exchange = contract_values['m_exchange']
        contract.m_secType = contract_values['m_secType']
        # If expiry is empty it will download all available expiries
        if contract_values['m_expiry'] <> "":
            contract.m_expiry = contract_values['m_expiry']
        self.connection.reqContractDetails(reqId, contract)
        sleep(20)
    
    def get_market_data(self):
        """
        Call for all the options prices and greeks
        """
        print "Calling Market Data"
        self.reqId = 1
        # Loop through all options contracts
        for option in self.d_contracts.values():
            self.d_ticker_reqId[self.reqId] = option
            self.connection.reqMktData(self.reqId, option, None, True)
            self.reqId += 1
            sleep(1)
        sleep(10)
        
    def save_option_contracts_to_dict(self, opt_contract):
        """
        It saves the options contracts downloaded as ContractDetails object
        to a dictionary of dictionaries
        """
        self.d_opt_contracts[opt_contract.m_summary.m_localSymbol]['m_conId']=opt_contract.m_summary.m_conId
        self.d_opt_contracts[opt_contract.m_summary.m_localSymbol]['m_symbol']=opt_contract.m_summary.m_symbol
        self.d_opt_contracts[opt_contract.m_summary.m_localSymbol]['m_secType']=opt_contract.m_summary.m_secType
        self.d_opt_contracts[opt_contract.m_summary.m_localSymbol]['m_expiry']=opt_contract.m_summary.m_expiry
        self.d_opt_contracts[opt_contract.m_summary.m_localSymbol]['m_strike']=opt_contract.m_summary.m_strike
        self.d_opt_contracts[opt_contract.m_summary.m_localSymbol]['m_right']=opt_contract.m_summary.m_right
        self.d_opt_contracts[opt_contract.m_summary.m_localSymbol]['m_multiplier']=opt_contract.m_summary.m_multiplier
        self.d_opt_contracts[opt_contract.m_summary.m_localSymbol]['m_exchange']=opt_contract.m_summary.m_exchange
        self.d_opt_contracts[opt_contract.m_summary.m_localSymbol]['m_currency']=opt_contract.m_summary.m_currency
        self.d_opt_contracts[opt_contract.m_summary.m_localSymbol]['m_localSymbol']=opt_contract.m_summary.m_localSymbol
        """
        It saves the options contracts downloaded as ContractDetails object
        to a dictionary of contracts
        """
        self.d_contracts[opt_contract.m_summary.m_localSymbol] = Contract()
        self.d_contracts[opt_contract.m_summary.m_localSymbol].m_conId=opt_contract.m_summary.m_conId
        self.d_contracts[opt_contract.m_summary.m_localSymbol].m_symbol=opt_contract.m_summary.m_symbol
        self.d_contracts[opt_contract.m_summary.m_localSymbol].m_secType=opt_contract.m_summary.m_secType
        self.d_contracts[opt_contract.m_summary.m_localSymbol].m_expiry=opt_contract.m_summary.m_expiry
        self.d_contracts[opt_contract.m_summary.m_localSymbol].m_strike=opt_contract.m_summary.m_strike
        self.d_contracts[opt_contract.m_summary.m_localSymbol].m_right=opt_contract.m_summary.m_right
        self.d_contracts[opt_contract.m_summary.m_localSymbol].m_multiplier=opt_contract.m_summary.m_multiplier
        self.d_contracts[opt_contract.m_summary.m_localSymbol].m_exchange=opt_contract.m_summary.m_exchange
        self.d_contracts[opt_contract.m_summary.m_localSymbol].m_currency=opt_contract.m_summary.m_currency
        self.d_contracts[opt_contract.m_summary.m_localSymbol].m_localSymbol=opt_contract.m_summary.m_localSymbol
        
    
    
 
if __name__ == '__main__':
    # Connection
    ib = IB_API()
    
    # Get contract details
    filename = excel.select_excel_file()
    xl = pd.ExcelFile(filename)
    df_input = xl.parse('input')
    #contract_values = {'m_symbol': 'TLT', 'm_exchange': 'SMART', 'm_secType': 'OPT', 
    #                  'm_expiry': '20140905'}
    
    contract_values = {
                       'm_symbol': str(df_input['m_symbol'][0]),
                       'm_exchange': str(df_input['m_exchange'][0]),
                       'm_secType': str(df_input['m_secType'][0]),
                       'm_expiry': str(df_input['m_expiry'][0])
                       }
    ib.get_contract_details(1, contract_values)
    
    
    

    
    # Get Market values
    ib.get_market_data()
    
    print ib.d_opt_contracts
    
    # Output
    columns_after_hours = [
                           'm_conId', 'm_localSymbol', 'm_symbol', 'm_currency', 'm_exchange', 'm_secType', 'm_multiplier', 'm_right', 'm_strike', 
                           'm_expiry', 'close', 'bid', 'ask'
                           ]
    columns_open_market = [
                           'bid_impliedVolatility', 'ask_impliedVolatility',  'bid_delta', 'ask_delta', 
                           'bid_theta', 'ask_theta', 'bid_gamma', 'ask_gamma', 'bid_pvDividend', 'ask_pvDividend', 'bid_vega', 'ask_vega', 
                           'bid_optPrice', 'ask_optPrice', 'bid_undPrice', 'ask_undPrice'     
                           ]
    df_option_chain = pd.DataFrame(ib.d_opt_contracts)
    df_option_chain = df_option_chain.T
    
    # Check if market is open
    if 'bid_impliedVolatility' in df_option_chain.columns:
        df_option_chain = df_option_chain[columns_after_hours + columns_open_market]
    else:
        df_option_chain = df_option_chain[columns_after_hours]
    
    # Save in a new excel tab
    excel.save_in_new_tab_of_excel(filename, df_option_chain, "option_chain")
    
    
    

 
 
 
 
 
 
 
 